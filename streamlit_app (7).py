"""
Aivora - Agent
Your Smarter Support for Every Student's Application
Streamlit App

Supabase-backed web app. UI restored to the previous Aivora design while keeping
Supabase as the backend and the local Selenium Worker as the processor.
"""

import io
import json
import os
import re
import time
import threading

import requests
from datetime import datetime, timezone

import bcrypt
import openpyxl
import pandas as pd
import streamlit as st
from cryptography.fernet import Fernet
from supabase import create_client, Client

try:
    import gspread
    from google.oauth2.service_account import Credentials
except Exception:
    gspread = None
    Credentials = None

ADMIN_USERNAME = st.secrets.get("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD", "admin123")

FINAL_STATUSES = {
    "مقبول نهائي", "قبول نهائي", "تم الرفض", "مرفوض نهائيًا",
    "مرفوض نهائيا", "مرفوض", "خالص",
}

@st.cache_resource(show_spinner=False)
def get_supabase() -> Client:
    url = st.secrets.get("SUPABASE_URL", os.getenv("SUPABASE_URL"))
    key = st.secrets.get("SUPABASE_SERVICE_ROLE_KEY", os.getenv("SUPABASE_SERVICE_ROLE_KEY"))
    if not url or not key:
        raise RuntimeError("مفيش SUPABASE_URL أو SUPABASE_SERVICE_ROLE_KEY في Secrets.")
    return create_client(url, key)

def db(): return get_supabase()
def now_iso(): return datetime.now(timezone.utc).isoformat()

def get_office_by_name(name):
    rows=db().table("offices").select("id,name,email,status,created_at").eq("name",str(name).strip()).limit(1).execute().data or []
    return rows[0] if rows else None

def get_office_with_hash(office_id):
    rows=db().table("offices").select("id,name,email,status,password_hash,created_at").eq("id",office_id).limit(1).execute().data or []
    return rows[0] if rows else None

def get_office_by_email(email):
    rows=db().table("offices").select("id,name,email,status").eq("email",str(email).strip().lower()).limit(1).execute().data or []
    return rows[0] if rows else None

def register_office(name,email,password):
    name,email,password=name.strip(),email.strip().lower(),password.strip()
    if not name or not email or not password:return False,"كمّلي كل البيانات الأول."
    if len(password)<6:return False,"الباسورد لازم يكون 6 أحرف على الأقل."
    try:
        if get_office_by_name(name) or get_office_by_email(email):return False,"اسم المكتب أو الإيميل مسجل قبل كده."
        password_hash=bcrypt.hashpw(password.encode(),bcrypt.gensalt()).decode()
        db().table("offices").insert({"name":name,"email":email,"password_hash":password_hash,"status":"pending"}).execute()
        return True,"تم التسجيل بنجاح. الحساب في انتظار موافقة الإدارة."
    except Exception as exc:return False,f"حصل خطأ أثناء التسجيل: {exc}"

def check_login(name,password):
    try:
        office=get_office_by_name(name)
        if not office:return False,"اسم المكتب أو الباسورد غلط!"
        if office.get("status")=="pending":return False,"حسابك في انتظار موافقة الإدارة!"
        if office.get("status")!="approved":return False,"حسابك موقوف!"
        full=get_office_with_hash(office["id"]); password_hash=str((full or {}).get("password_hash") or "")
        if not password_hash or not bcrypt.checkpw(password.encode(),password_hash.encode()):return False,"اسم المكتب أو الباسورد غلط!"
        return True,office
    except Exception as exc:return False,f"حصل خطأ في تسجيل الدخول: {exc}"

def get_pending_accounts():
    try:return db().table("offices").select("id,name,email,status,created_at").eq("status","pending").order("created_at").execute().data or []
    except Exception:return []

def set_office_status(office_id,status):db().table("offices").update({"status":status}).eq("id",office_id).execute()
def log_activity(office_id,action,file_name="",details=None,student_id=None,data_source_id=None):
    try:db().table("activity_logs").insert({"office_id":office_id,"student_record_id":student_id,"data_source_id":data_source_id,"action":action,"file_name":file_name,"details":details or {}}).execute()
    except Exception as exc:print(f"Activity log error: {exc}")

def get_gsheet_client():
    if gspread is None or Credentials is None:raise RuntimeError("مكتبات Google Sheets مش مثبتة.")
    creds=Credentials.from_service_account_info(st.secrets["gcp_service_account"],scopes=["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/drive"])
    return gspread.authorize(creds)
def extract_sheet_id(link):
    m=re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)",str(link));return m.group(1) if m else None
def extract_gid(link):
    m=re.search(r"[?#&]gid=(\d+)",str(link));return int(m.group(1)) if m else None
def read_gsheet_rows(link):
    sid=extract_sheet_id(link)
    if not sid:raise ValueError("رابط Google Sheets غير صحيح.")
    ss=get_gsheet_client().open_by_key(sid);gid=extract_gid(link);ws=next((w for w in ss.worksheets() if w.id==gid),ss.sheet1) if gid is not None else ss.sheet1
    return ws.get_all_values()
def get_saved_gsheet_link(office_id):
    rows=db().table("data_sources").select("source_url,created_at").eq("office_id",office_id).eq("source_type","google_sheet").not_.is_("source_url","null").order("created_at",desc=True).limit(1).execute().data or []
    return rows[0].get("source_url") if rows else None
def save_gsheet_link(office_id,link):
    try:
        existing=db().table("data_sources").select("id").eq("office_id",office_id).eq("source_type","google_sheet").eq("source_url",link).limit(1).execute().data or []
        if existing:return True,"تم حفظ الرابط بالفعل"
        db().table("data_sources").insert({"office_id":office_id,"source_type":"google_sheet","source_name":"Google Sheet","source_url":link,"column_mapping":{}}).execute();return True,"تم حفظ الرابط بنجاح"
    except Exception as exc:return False,str(exc)

def find_excel_columns(ws):
    cols={"name":None,"email":None,"password":None};header_row=None
    for row_idx,row in enumerate(ws.iter_rows(min_row=1,max_row=5,values_only=True),start=1):
        values=[str(c).strip() if c is not None else "" for c in row]
        if any("يميل" in v or "mail" in v.lower() or "بريد" in v for v in values):
            header_row=row_idx
            for i,cell in enumerate(values):
                low=cell.lower()
                if any(k in cell for k in ["اسم","الإسم","الاسم"]) or "name" in low:cols["name"]=i
                elif any(k in cell for k in ["يميل","بريد"]) or "mail" in low:cols["email"]=i
                elif any(k in cell for k in ["باسورد","كلمة المرور","password","pass"]):cols["password"]=i
            break
    if header_row is None:raise ValueError("مش لاقي هيدر الإكسيل.")
    if cols["email"] is None:raise ValueError("مش لاقي عمود الإيميل.")
    if cols["password"] is None:raise ValueError("مش لاقي عمود الباسورد.")
    if cols["name"] is None:cols["name"]=cols["email"]
    return cols,header_row

def parse_excel_bytes(file_bytes):
    wb=openpyxl.load_workbook(io.BytesIO(file_bytes),data_only=False);ws=wb.active;cols,header_row=find_excel_columns(ws);records=[]
    for excel_row,row in enumerate(ws.iter_rows(min_row=header_row+1,values_only=True),start=header_row+1):
        values=list(row);email=str(values[cols["email"]] or "").strip() if cols["email"]<len(values) else "";password=str(values[cols["password"]] or "").strip() if cols["password"]<len(values) else "";name=str(values[cols["name"]] or "").strip() if cols["name"]<len(values) else email
        if not email or not password:continue
        records.append({"source_row_number":excel_row,"student_name":name or email,"login_identifier":email,"password":password,"original_data":{f"column_{i+1}":(str(v) if v is not None else "") for i,v in enumerate(values)}})
    return records

def encrypt_password(password):
    key=st.secrets.get("STUDENT_PASSWORD_ENCRYPTION_KEY",os.getenv("STUDENT_PASSWORD_ENCRYPTION_KEY"))
    if not key:raise RuntimeError("STUDENT_PASSWORD_ENCRYPTION_KEY مش موجود في Secrets.")
    return Fernet(key.encode()).encrypt(password.encode()).decode()
def upsert_students(office_id,source_id,records):
    payload=[{"office_id":office_id,"data_source_id":source_id,"source_row_number":r["source_row_number"],"student_name":r["student_name"],"login_identifier":r["login_identifier"],"encrypted_password":encrypt_password(r["password"]),"application_status":"","original_data":r["original_data"],"updated_at":now_iso()} for r in records]
    db().table("student_records").insert(payload).execute();return len(payload)
def create_data_source(office_id,source_type,source_name,source_url=None,mapping=None):
    source_type={"xlsx":"excel","xls":"excel","gsheet":"google_sheet","google_sheet":"google_sheet","csv":"csv"}.get(source_type,source_type)
    row=db().table("data_sources").insert({"office_id":office_id,"source_type":source_type,"source_name":source_name,"source_url":source_url,"column_mapping":mapping or {}}).execute().data;return row[0]
def import_students(office_id,source_type,source_name,file_bytes=None,source_url=None):
    if source_type in {"xlsx","xls","excel"}:records=parse_excel_bytes(file_bytes)
    else:
        rows=read_gsheet_rows(source_url)
        if not rows:raise ValueError("الشيت فاضي.")
        wb=openpyxl.Workbook();ws=wb.active
        for row in rows:ws.append(row)
        out=io.BytesIO();wb.save(out);records=parse_excel_bytes(out.getvalue())
    if not records:raise ValueError("مش لاقي طلاب عندهم إيميل وباسورد صالحين.")
    source=create_data_source(office_id,source_type,source_name,source_url=source_url);count=upsert_students(office_id,source["id"],records);return source,count

def create_job(office_id,source,file_name):
    rows=db().table("jobs").insert({"office_id":office_id,"data_source_id":source["id"],"source_type":source["source_type"],"source_ref":source.get("source_url") or source["id"],"file_name":file_name,"status":"pending"}).execute().data;return rows[0]
def get_job(job_id):
    rows=db().table("jobs").select("*").eq("id",job_id).limit(1).execute().data or [];return rows[0] if rows else None
def get_job_progress_rows(job_id):return db().table("job_progress").select("student_index,total,student_name,status,created_at").eq("job_id",job_id).order("student_index").execute().data or []
def get_students(office_id,search=""):
    rows=db().table("student_records").select("id,student_name,login_identifier,application_status,status_updated_at,source_row_number,created_at,updated_at").eq("office_id",office_id).execute().data or []
    def ts(v):
        try:return datetime.fromisoformat(str(v or "").replace("Z","+00:00"))
        except Exception:return datetime.min.replace(tzinfo=timezone.utc)
    latest={}
    for row in rows:
        key=str(row.get("login_identifier") or row.get("student_name") or "").strip().lower()
        if not key:continue
        score=(ts(row.get("status_updated_at")),ts(row.get("updated_at")),ts(row.get("created_at")))
        if key not in latest or score>latest[key][0]:latest[key]=(score,row)
    rows=sorted([v[1] for v in latest.values()],key=lambda r:str(r.get("student_name") or "").lower())
    if search.strip():
        q=search.strip().lower();rows=[r for r in rows if q in str(r.get("student_name","")).lower()]
    return rows

def _decrypt_student_password(value):
    key=st.secrets.get("STUDENT_PASSWORD_ENCRYPTION_KEY",os.getenv("STUDENT_PASSWORD_ENCRYPTION_KEY"))
    if not key:raise RuntimeError("STUDENT_PASSWORD_ENCRYPTION_KEY مش موجود في Secrets.")
    return Fernet(key.encode()).decrypt(str(value).encode()).decode()
BASE_URL="https://apiadm.study-in-egypt.gov.eg/api";SITE_URL="https://admission.study-in-egypt.gov.eg";WORKER_WAIT_SECONDS=30

def _legacy_api_login(email,password):
    session=requests.Session();session.headers.update({"accept":"application/json, text/plain, */*","accept-language":"ar","device":"CITIZEN","origin":SITE_URL,"referer":SITE_URL+"/","user-agent":"Mozilla/5.0","content-type":"application/json"})
    response=session.post(f"{BASE_URL}/student/login",json={"email":email,"password":password},timeout=30)
    if response.status_code not in (200,201):return None,None,f"فشل تسجيل الدخول - كود: {response.status_code}"
    body=response.json() if response.content else {};return session,body.get("token","") or response.headers.get("x-csrf-token",""),None

def _legacy_api_get_status(session,token):
    filt={"where":{},"limit":10,"offset":0,"order":"statusUpdatedAt DESC","fields":["serviceSlug","ID","createdAt","statusUpdatedAt","activityId","activityName"]};headers={"x-csrf-token":token} if token else {}
    response=session.get(f"{BASE_URL}/dynamic_services/inbox",params={"filter":json.dumps(filt)},headers=headers,timeout=30)
    if response.status_code not in (200,304):return f"خطأ في جلب الحالة ({response.status_code})"
    result=response.json().get("result") or []
    if not result:return "مفيش طلبات"
    activity=result[0].get("activityName") or "غير محدد";mapping={"قبول الفحص الفنى":"القبول المبدئي","قبول الفحص الفني":"القبول المبدئي","تم السداد":"تم السداد","تأكيد استلام الملف وصحة و اكتمال المستندات":"تأكيد استلام الملف وصحة واكتمال المستندات","الانتظار مراجعة الطلب":"بانتظار مراجعة الطلب","قبول من رئيس الادارة المركزية":"قبول من رئيس الإدارة المركزية"};return mapping.get(activity,activity)
def _legacy_api_logout(session):
    if session is None:return
    try:session.post(f"{BASE_URL}/student/logout",json={"redirectUrl":SITE_URL},timeout=15)
    except Exception:pass
def _claim_fallback_job(job_id):
    rows=db().table("jobs").update({"status":"processing","started_at":now_iso(),"claimed_by":"streamlit-fallback"}).eq("id",job_id).eq("status","pending").select("*").execute().data or [];return rows[0] if rows else None
def _run_legacy_api_fallback(job):
    students=db().table("student_records").select("*").eq("office_id",job["office_id"]).eq("data_source_id",job["data_source_id"]).order("source_row_number").execute().data or []
    students=[r for r in students if str(r.get("application_status") or "").strip() not in FINAL_STATUSES and r.get("login_identifier") and r.get("encrypted_password")]
    total=len(students)
    if not total:db().table("jobs").update({"status":"done","finished_at":now_iso()}).eq("id",job["id"]).execute();return
    technical=[]
    for index,student in enumerate(students,1):
        session=None;status="خطأ فني في الفحص";is_tech=False
        try:
            session,token,error=_legacy_api_login(student["login_identifier"],_decrypt_student_password(student["encrypted_password"]))
            if error:status="فشل تسجيل الدخول"
            else:
                status=_legacy_api_get_status(session,token);is_tech=str(status).startswith("خطأ")
                if is_tech:status="خطأ فني في الفحص"
        except Exception as exc:is_tech=True;technical.append((student,exc));status="خطأ فني في الفحص"
        finally:_legacy_api_logout(session)
        stamp=now_iso();db().table("student_records").update({"application_status":status,"status_updated_at":stamp,"updated_at":stamp}).eq("id",student["id"]).execute()
        db().table("job_progress").insert({"job_id":job["id"],"student_index":index,"total":total,"student_name":student.get("student_name") or student.get("login_identifier"),"status":status}).execute()
    for retry_index,(student,_) in enumerate(technical,1):
        session=None;status="خطأ فني في الفحص"
        try:
            session,token,error=_legacy_api_login(student["login_identifier"],_decrypt_student_password(student["encrypted_password"]))
            if not error:status=_legacy_api_get_status(session,token)
        except Exception:pass
        finally:_legacy_api_logout(session)
        stamp=now_iso();db().table("student_records").update({"application_status":status,"status_updated_at":stamp,"updated_at":stamp}).eq("id",student["id"]).execute();db().table("job_progress").insert({"job_id":job["id"],"student_index":total+retry_index,"total":total,"student_name":student.get("student_name") or student.get("login_identifier"),"status":status}).execute()
    db().table("jobs").update({"status":"done","finished_at":now_iso()}).eq("id",job["id"]).execute()
def wait_for_worker_or_legacy_fallback(job_id):
    deadline=time.monotonic()+WORKER_WAIT_SECONDS
    while time.monotonic()<deadline:
        job=get_job(job_id)
        if not job:return "missing"
        if str(job.get("status") or "pending")!="pending":return "worker"
        time.sleep(2)
    claimed=_claim_fallback_job(job_id)
    if claimed:_run_legacy_api_fallback(claimed);return "fallback"
    return "worker"
def _background_update_job(job_id):
    try:wait_for_worker_or_legacy_fallback(job_id)
    except Exception as exc:
        try:db().table("jobs").update({"status":"failed","finished_at":now_iso(),"error":str(exc)[:1000]}).eq("id",job_id).execute()
        except Exception as db_exc:print(f"Background job save error: {db_exc}")

st.set_page_config(page_title="Aivora - Agent",page_icon="✨",layout="wide",initial_sidebar_state="collapsed")
import streamlit.components.v1 as components
components.html("""<script>(function(){try{var d=window.parent.document;if(!d.querySelector('meta[name=\"google\"][content=\"notranslate\"]')){var m=d.createElement('meta');m.name='google';m.content='notranslate';d.head.appendChild(m);}d.documentElement.classList.add('notranslate');d.documentElement.setAttribute('translate','no');}catch(e){}})();</script>""",height=0,width=0)
st.markdown("""<style>
html,body,[class*="css"],.stApp{font-family:'Cairo',sans-serif!important;direction:rtl}.stApp{background:#f5f7fb;color:#111827}.block-container{max-width:1180px;padding-top:1.4rem;padding-bottom:3rem}#MainMenu,footer,header,[data-testid="stToolbar"],[data-testid="stDecoration"],[data-testid="stStatusWidget"],[data-testid="stSidebarNav"]{display:none!important}.card{background:#fff;border:1px solid #e5e7eb;border-radius:18px;padding:22px;box-shadow:0 3px 14px rgba(17,24,39,.045);margin-bottom:18px}.topbar{background:#fff;border:1px solid #e5e7eb;border-radius:16px;padding:13px 18px;margin-bottom:22px}.brand-title{font-size:18px;font-weight:800;color:#111827}.brand-sub{font-size:12px;color:#6b7280}.section-title{font-size:18px;font-weight:800;color:#111827;margin:5px 0 13px}.section-sub{color:#6b7280;font-size:13px;margin-bottom:14px}.stButton>button{width:100%;min-height:44px;border-radius:10px!important;background:#2563eb!important;color:#fff!important;font-weight:700!important}.success-box{background:#ecfdf5;border:1px solid #bbf7d0;border-radius:14px;padding:16px}.success-title{font-size:17px;font-weight:800;color:#166534}.status-badge{display:inline-block;padding:4px 10px;border-radius:999px;font-size:12px;font-weight:700}.status-ok{background:#ecfdf5;color:#15803d}.status-warn{background:#fffbeb;color:#b45309}.status-error{background:#fef2f2;color:#b91c1c}.status-info{background:#eff6ff;color:#1d4ed8}.connected-box{background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;padding:11px 14px;color:#166534;font-size:13px}[data-testid="InputInstructions"]{display:none!important}[data-testid="stTextInput"] button{display:none!important}
</style>""",unsafe_allow_html=True)
for key,default in [("logged_in",False),("is_admin",False),("office",None),("update_locked",False),("pending_file_bytes",None),("pending_filename",""),("active_job_id",None),("update_start_requested",False),("update_preparing",False)]:
    if key not in st.session_state:st.session_state[key]=default
if not st.session_state.logged_in and not st.session_state.is_admin:
    st.markdown("<div style='text-align:center;margin:42px 0 24px;'><div style='font-size:48px;'>✨</div><div style='font-size:31px;font-weight:800;color:#111827;'>Aivora</div><div style='font-size:14px;color:#6b7280;'>Your Smarter Support for Every Student's Application</div></div>",unsafe_allow_html=True)
    left,right=st.columns([1.15,1],gap="large")
    with left:st.markdown("<div style='padding:40px 20px;'><div style='font-size:14px;color:#2563eb;font-weight:800;'>حل بسيط لإدارة المتابعة</div><div style='font-size:31px;font-weight:800;color:#111827;'>تابع طلبات طلابك<br>من مكان واحد.</div><div style='font-size:15px;color:#6b7280;line-height:1.9;margin-top:12px;'>حدّث حالات الطلبات، اربط Google Sheets، وابحث عن أي طالب بسرعة بدون متابعة يدوية.</div></div>",unsafe_allow_html=True)
    with right:
        st.markdown('<div class="card">',unsafe_allow_html=True);tab1,tab2=st.tabs(["تسجيل الدخول","حساب جديد"])
        with tab1:
            username=st.text_input("اسم المكتب",key="login_user");password=st.text_input("كلمة المرور",type="password",key="login_pass")
            if st.button("تسجيل الدخول",key="login_btn"):
                if username==ADMIN_USERNAME and password==ADMIN_PASSWORD:st.session_state.is_admin=True;st.rerun()
                else:
                    ok,result=check_login(username,password)
                    if ok:st.session_state.logged_in=True;st.session_state.office=result;st.session_state.update_locked=False;st.rerun()
                    else:st.error(result)
        with tab2:
            new_office=st.text_input("اسم المكتب",key="reg_office");new_email=st.text_input("الإيميل",key="reg_email");new_pass=st.text_input("كلمة المرور",type="password",key="reg_pass");new_pass2=st.text_input("تأكيد كلمة المرور",type="password",key="reg_pass2")
            if st.button("إنشاء الحساب",key="reg_btn"):
                if not new_office or not new_email or not new_pass:st.error("املأ كل الحقول!")
                elif new_pass!=new_pass2:st.error("كلمة المرور مش متطابقة!")
                else:
                    ok,msg=register_office(new_office,new_email,new_pass);(st.success if ok else st.error)(msg)
        st.markdown('</div>',unsafe_allow_html=True)
    st.stop()
if st.session_state.is_admin:
    st.markdown("<div class='card'><div class='section-title'>الإدارة</div></div>",unsafe_allow_html=True);pending=get_pending_accounts()
    for acc in pending:
        c1,c2,c3,c4=st.columns([3,3,1,1]);c1.write(acc.get("name",""));c2.write(acc.get("email",""))
        if c3.button("قبول",key=f"approve_{acc['id']}"):set_office_status(acc["id"],"approved");st.rerun()
        if c4.button("رفض",key=f"reject_{acc['id']}"):set_office_status(acc["id"],"rejected");st.rerun()
    if st.button("تسجيل الخروج من الإدارة",key="admin_logout"):st.session_state.clear();st.rerun()
    st.stop()
office=st.session_state.office
if not office:st.session_state.logged_in=False;st.rerun()
office_id=office["id"]
st.markdown("<div class='topbar'><div class='brand-title'>Aivora</div><div class='brand-sub'>Your Smarter Support for Every Student's Application</div></div>",unsafe_allow_html=True)
st.markdown('<div class="card">',unsafe_allow_html=True);st.markdown('<div class="section-title">مصدر بيانات الطلاب</div><div class="section-sub">اختر الطريقة التي يحتوي بها ملف الطلاب.</div>',unsafe_allow_html=True)
source_options=["📂 رفع ملف Excel","🔗 ربط Google Sheets"];source=st.radio("",source_options,horizontal=True,label_visibility="collapsed",key="source_mode");file_bytes=None;filename="";saved_link=get_saved_gsheet_link(office_id);sheet_id_source=None
if source=="📂 رفع ملف Excel":
    uploaded=st.file_uploader("ارفع ملف Excel",type=["xlsx","xls","csv"],label_visibility="collapsed",key="excel_upload")
    if uploaded:file_bytes=uploaded.getvalue();filename=uploaded.name;st.success(f"تم اختيار الملف: {uploaded.name}")
else:
    if saved_link:
        st.markdown('<div class="connected-box">✓ Google Sheets متصل بالفعل لهذا المكتب</div>',unsafe_allow_html=True);change=st.checkbox("تغيير رابط الشيت",key="change_link")
        if change:
            new_link=st.text_input("رابط Google Sheets",key="new_link")
            if st.button("حفظ الرابط",key="save_link"):
                sid=extract_sheet_id(new_link) if new_link else None
                if not new_link:st.error("أدخل الرابط أولاً!")
                elif not sid:st.error("الرابط غير صحيح!")
                else:
                    ok,msg=save_gsheet_link(office_id,new_link);st.success("تم حفظ الرابط!") if ok else st.error(msg)
                    if ok:st.rerun()
        sheet_id_source=extract_sheet_id(saved_link)
    else:
        new_link=st.text_input("رابط Google Sheets",key="first_link")
        if st.button("حفظ وربط الشيت",key="save_first_link"):
            sid=extract_sheet_id(new_link) if new_link else None
            if sid:
                ok,msg=save_gsheet_link(office_id,new_link);st.success("تم حفظ الرابط بنجاح!") if ok else st.error(msg)
                if ok:st.rerun()
            else:st.error("الرابط غير صحيح!")
    if sheet_id_source:
        if st.button("تحميل بيانات الشيت",key="load_sheet"):
            rows=read_gsheet_rows(saved_link);wb=openpyxl.Workbook();ws=wb.active
            for r in rows:ws.append(r)
            out=io.BytesIO();wb.save(out);st.session_state.pending_file_bytes=out.getvalue();st.session_state.pending_filename="google_sheet";st.success("تم جلب بيانات الشيت. البيانات جاهزة للتحديث.")
        file_bytes=st.session_state.pending_file_bytes;filename=st.session_state.pending_filename
st.markdown('</div>',unsafe_allow_html=True)

def _request_update_start():st.session_state.update_start_requested=True;st.session_state.update_locked=True
def _prepare_update_job(office_id,source_type,source_name,file_bytes,source_url):
    try:
        src,count=import_students(office_id,source_type,source_name,file_bytes=file_bytes,source_url=source_url);job=create_job(office_id,src,source_name);t=threading.Thread(target=_background_update_job,args=(job["id"],),daemon=True);t.start();return job["id"],None
    except Exception as exc:return None,str(exc)
@st.fragment(run_every=2)
def render_processing():
    if not file_bytes and not st.session_state.active_job_id and not st.session_state.update_start_requested:return
    st.markdown('<div class="card">',unsafe_allow_html=True);st.markdown('<div class="section-title">تحديث حالات الطلاب</div><div class="section-sub">اضغط الزر لبدء فحص الطلبات وتحديث النتائج.</div>',unsafe_allow_html=True)
    if file_bytes and not st.session_state.update_locked and not st.session_state.active_job_id:st.button("▶ تحديث حالات الطلاب",key="start_main",on_click=_request_update_start)
    if st.session_state.update_start_requested and not st.session_state.active_job_id and not st.session_state.update_preparing:
        running=db().table("jobs").select("id").eq("office_id",office_id).in_("status",["pending","processing"]).limit(1).execute().data or []
        if running:st.session_state.update_start_requested=False;st.warning("في تحديث شغال بالفعل لهذا المكتب. استني لحد ما يخلص.")
        else:
            st.session_state.update_preparing=True;is_gsheet_source=bool(saved_link and source=="🔗 ربط Google Sheets");source_type="gsheet" if is_gsheet_source else "xlsx";source_name="Google Sheet" if is_gsheet_source else (filename or "students.xlsx");captured=file_bytes
            def runner():
                job_id,error=_prepare_update_job(office_id,source_type,source_name,captured,saved_link if is_gsheet_source else None)
                if job_id:st.session_state.active_job_id=job_id
                if error:print(f"Update preparation failed: {error}")
            threading.Thread(target=runner,daemon=True).start()
    job=get_job(st.session_state.active_job_id) if st.session_state.active_job_id else None
    if st.session_state.update_start_requested or st.session_state.update_preparing:
        if not job:st.info("⏳ سيبدأ التحديث خلال ثواني…")
    if job:
        status=str(job.get("status") or "pending")
        if status=="pending":st.info("⏳ سيبدأ التحديث خلال ثواني…")
        elif status=="processing":st.info("▶️ بدأ التحديث. جاري فحص الطلاب وتحديث الحالات…")
        rows=get_job_progress_rows(job["id"])
        if rows:
            latest={}
            for r in rows:
                k=str(r.get("student_name") or "").strip().lower()
                if k:latest[k]=r
            shown=list(latest.values());total=int(rows[-1].get("total") or 0);st.progress(min(len(shown)/max(total,1),1.0));st.caption(f"طالب {min(len(shown),total) if total else len(shown)} من {total}")
            last=rows[-1];st.info(f"🔄 آخر طالب تم فحصه: **{last.get('student_name') or 'طالب'}** — الحالة الجديدة: **{last.get('status') or ''}**")
            st.dataframe(pd.DataFrame([{"#":r.get("student_index"),"اسم الطالب":r.get("student_name","") ,"الحالة":r.get("status","")} for r in reversed(shown)]),use_container_width=True,hide_index=True)
        if status=="done":st.markdown('<div class="success-box"><div class="success-title">اكتمل التحديث 🎉</div></div>',unsafe_allow_html=True)
        elif status=="failed":st.info("تعذر إكمال التحديث حاليًا. حاولي مرة تانية بعد قليل.")
    elif st.session_state.update_locked:st.info("🔒 تم تشغيل تحديث بالفعل في هذه الجلسة. لو عايزة تبدئي تحديث جديد، سجّلي خروج وادخلي تاني.")
    st.markdown('</div>',unsafe_allow_html=True)
render_processing()
st.markdown('<div class="card">',unsafe_allow_html=True);st.markdown('<div class="section-title">البحث عن طالب</div>',unsafe_allow_html=True);search_query=st.text_input("اسم الطالب",label_visibility="collapsed",key="student_search")
if search_query:
    found=get_students(office_id,search_query)
    if found:
        for s in found:st.markdown(f'<div class="card"><b>👤 {s.get("student_name","")}</b><br><span class="status-badge status-info">{s.get("application_status") or "لم يتم الفحص بعد"}</span></div>',unsafe_allow_html=True)
    else:st.info("مفيش طالب بالاسم ده.")
st.markdown('</div>',unsafe_allow_html=True)
if st.button("تسجيل الخروج",key="logout_main"):st.session_state.clear();st.rerun()
