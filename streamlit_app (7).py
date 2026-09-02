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

def db():
    return get_supabase()

def now_iso():
    return datetime.now(timezone.utc).isoformat()

# ==================== Accounts ====================
def get_office_by_name(name):
    rows = db().table("offices").select("id,name,email,status,created_at").eq("name", str(name).strip()).limit(1).execute().data or []
    return rows[0] if rows else None

def get_office_with_hash(office_id):
    rows = db().table("offices").select("id,name,email,status,password_hash,created_at").eq("id", office_id).limit(1).execute().data or []
    return rows[0] if rows else None

def get_office_by_email(email):
    rows = db().table("offices").select("id,name,email,status").eq("email", str(email).strip().lower()).limit(1).execute().data or []
    return rows[0] if rows else None

def register_office(name, email, password):
    name, email, password = name.strip(), email.strip().lower(), password.strip()
    if not name or not email or not password: return False, "كمّلي كل البيانات الأول."
    if len(password) < 6: return False, "الباسورد لازم يكون 6 أحرف على الأقل."
    try:
        if get_office_by_name(name) or get_office_by_email(email): return False, "اسم المكتب أو الإيميل مسجل قبل كده."
        password_hash = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
        db().table("offices").insert({"name": name, "email": email, "password_hash": password_hash, "status": "pending"}).execute()
        return True, "تم التسجيل بنجاح. الحساب في انتظار موافقة الإدارة."
    except Exception as exc:
        return False, f"حصل خطأ أثناء التسجيل: {exc}"

def check_login(name, password):
    try:
        office = get_office_by_name(name)
        if not office: return False, "اسم المكتب أو الباسورد غلط!"
        if office.get("status") == "pending": return False, "حسابك في انتظار موافقة الإدارة!"
        if office.get("status") != "approved": return False, "حسابك موقوف!"
        full = get_office_with_hash(office["id"])
        password_hash = str((full or {}).get("password_hash") or "")
        if not password_hash or not bcrypt.checkpw(password.encode(), password_hash.encode()): return False, "اسم المكتب أو الباسورد غلط!"
        return True, office
    except Exception as exc:
        return False, f"حصل خطأ في تسجيل الدخول: {exc}"

def get_pending_accounts():
    try:
        return db().table("offices").select("id,name,email,status,created_at").eq("status", "pending").order("created_at").execute().data or []
    except Exception: return []

def set_office_status(office_id, status):
    db().table("offices").update({"status": status}).eq("id", office_id).execute()

def log_activity(office_id, action, file_name="", details=None, student_id=None, data_source_id=None):
    try:
        db().table("activity_logs").insert({"office_id": office_id, "student_record_id": student_id, "data_source_id": data_source_id, "action": action, "file_name": file_name, "details": details or {}}).execute()
    except Exception as exc: print(f"Activity log error: {exc}")

# ==================== Google Sheet ====================
def get_gsheet_client():
    if gspread is None or Credentials is None: raise RuntimeError("مكتبات Google Sheets مش مثبتة.")
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"])
    return gspread.authorize(creds)

def extract_sheet_id(link):
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", str(link)); return m.group(1) if m else None

def extract_gid(link):
    m = re.search(r"[?#&]gid=(\d+)", str(link)); return int(m.group(1)) if m else None

def read_gsheet_rows(link):
    sid = extract_sheet_id(link)
    if not sid: raise ValueError("رابط Google Sheets غير صحيح.")
    ss = get_gsheet_client().open_by_key(sid)
    gid = extract_gid(link)
    ws = next((w for w in ss.worksheets() if w.id == gid), ss.sheet1) if gid is not None else ss.sheet1
    return ws.get_all_values()

def get_saved_gsheet_link(office_id):
    rows = db().table("data_sources").select("source_url,created_at").eq("office_id", office_id).eq("source_type", "google_sheet").not_.is_("source_url", "null").order("created_at", desc=True).limit(1).execute().data or []
    return rows[0].get("source_url") if rows else None

def save_gsheet_link(office_id, link):
    try:
        existing = db().table("data_sources").select("id").eq("office_id", office_id).eq("source_type", "google_sheet").eq("source_url", link).limit(1).execute().data or []
        if existing: return True, "تم حفظ الرابط بالفعل"
        db().table("data_sources").insert({"office_id": office_id, "source_type": "google_sheet", "source_name": "Google Sheet", "source_url": link, "column_mapping": {}}).execute()
        return True, "تم حفظ الرابط بنجاح"
    except Exception as exc: return False, str(exc)

# ==================== Excel / student import ====================
def find_excel_columns(ws):
    cols = {"name": None, "email": None, "password": None}; header_row = None
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), start=1):
        values = [str(c).strip() if c is not None else "" for c in row]
        if any("يميل" in v or "mail" in v.lower() or "بريد" in v for v in values):
            header_row = row_idx
            for i, cell in enumerate(values):
                low = cell.lower()
                if any(k in cell for k in ["اسم", "الإسم", "الاسم"]) or "name" in low: cols["name"] = i
                elif any(k in cell for k in ["يميل", "بريد"]) or "mail" in low: cols["email"] = i
                elif any(k in cell for k in ["باسورد", "كلمة المرور", "password", "pass"]): cols["password"] = i
            break
    if header_row is None: raise ValueError("مش لاقي هيدر الإكسيل.")
    if cols["email"] is None: raise ValueError("مش لاقي عمود الإيميل.")
    if cols["password"] is None: raise ValueError("مش لاقي عمود الباسورد.")
    if cols["name"] is None: cols["name"] = cols["email"]
    return cols, header_row

def parse_excel_bytes(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False); ws = wb.active; cols, header_row = find_excel_columns(ws); records=[]
    for excel_row, row in enumerate(ws.iter_rows(min_row=header_row+1, values_only=True), start=header_row+1):
        values=list(row); email=str(values[cols["email"]] or "").strip() if cols["email"] < len(values) else ""; password=str(values[cols["password"]] or "").strip() if cols["password"] < len(values) else ""; name=str(values[cols["name"]] or "").strip() if cols["name"] < len(values) else email
        if not email or not password: continue
        records.append({"source_row_number":excel_row,"student_name":name or email,"login_identifier":email,"password":password,"original_data":{f"column_{i+1}":(str(v) if v is not None else "") for i,v in enumerate(values)}})
    return records

def encrypt_password(password):
    key=st.secrets.get("STUDENT_PASSWORD_ENCRYPTION_KEY", os.getenv("STUDENT_PASSWORD_ENCRYPTION_KEY"))
    if not key: raise RuntimeError("STUDENT_PASSWORD_ENCRYPTION_KEY مش موجود في Secrets.")
    return Fernet(key.encode()).encrypt(password.encode()).decode()

def upsert_students(office_id, source_id, records):
    payload=[{"office_id":office_id,"data_source_id":source_id,"source_row_number":r["source_row_number"],"student_name":r["student_name"],"login_identifier":r["login_identifier"],"encrypted_password":encrypt_password(r["password"]),"application_status":"","original_data":r["original_data"],"updated_at":now_iso()} for r in records]
    db().table("student_records").insert(payload).execute(); return len(payload)

def create_data_source(office_id, source_type, source_name, source_url=None, mapping=None):
    source_type={"xlsx":"excel","xls":"excel","gsheet":"google_sheet","google_sheet":"google_sheet","csv":"csv"}.get(source_type, source_type)
    row=db().table("data_sources").insert({"office_id":office_id,"source_type":source_type,"source_name":source_name,"source_url":source_url,"column_mapping":mapping or {}}).execute().data
    return row[0]

def import_students(office_id, source_type, source_name, file_bytes=None, source_url=None):
    if source_type in {"xlsx","xls","excel"}: records=parse_excel_bytes(file_bytes)
    else:
        rows=read_gsheet_rows(source_url)
        if not rows: raise ValueError("الشيت فاضي.")
        wb=openpyxl.Workbook(); ws=wb.active
        for row in rows: ws.append(row)
        out=io.BytesIO(); wb.save(out); records=parse_excel_bytes(out.getvalue())
    if not records: raise ValueError("مش لاقي طلاب عندهم إيميل وباسورد صالحين.")
    source=create_data_source(office_id, source_type, source_name, source_url=source_url); count=upsert_students(office_id, source["id"], records)
    log_activity(office_id,"استيراد بيانات الطلاب إلى Supabase",source_name,{"students":count,"source_type":source["source_type"]},data_source_id=source["id"])
    return source,count

# ==================== Jobs ====================
def create_job(office_id, source, file_name):
    rows=db().table("jobs").insert({"office_id":office_id,"data_source_id":source["id"],"source_type":source["source_type"],"source_ref":source.get("source_url") or source["id"],"file_name":file_name,"status":"pending"}).execute().data
    return rows[0]

def get_job(job_id):
    rows=db().table("jobs").select("*").eq("id",job_id).limit(1).execute().data or []; return rows[0] if rows else None

def get_job_progress(job_id):
    return db().table("job_progress").select("student_index,total,student_name,status,created_at").eq("job_id",job_id).order("student_index", desc=True).limit(1).execute().data or []

def get_job_progress_rows(job_id):
    return db().table("job_progress").select("student_index,total,student_name,status,created_at").eq("job_id",job_id).order("student_index").execute().data or []

def get_students(office_id, search=""):
    rows=(
        db().table("student_records")
        .select("id,student_name,login_identifier,application_status,status_updated_at,source_row_number,created_at,updated_at")
        .eq("office_id",office_id)
        .execute().data or []
    )

    def _ts(value):
        text=str(value or "").strip()
        if not text:
            return datetime.min.replace(tzinfo=timezone.utc)
        try:
            return datetime.fromisoformat(text.replace("Z","+00:00"))
        except Exception:
            return datetime.min.replace(tzinfo=timezone.utc)

    latest={}
    for row in rows:
        key=str(row.get("login_identifier") or row.get("student_name") or "").strip().lower()
        if not key:
            continue
        score=(_ts(row.get("status_updated_at")),_ts(row.get("updated_at")),_ts(row.get("created_at")))
        current=latest.get(key)
        if current is None or score > current[0]:
            latest[key]=(score,row)
    rows=[item[1] for item in latest.values()]
    rows=sorted(rows,key=lambda r:str(r.get("student_name") or "").lower())
    if search.strip():
        q=search.strip().lower()
        rows=[r for r in rows if q in str(r.get("student_name","")).lower()]
    return rows

def status_class(status):
    s=str(status or "")
    if any(x in s for x in ["مقبول","خالص"]): return "status-ok"
    if any(x in s for x in ["رفض","فشل","خطأ"]): return "status-error"
    if any(x in s for x in ["مفيش","انتظار","مراجعة"]): return "status-warn"
    return "status-info"



# ==================== LEGACY API FALLBACK 30S ====================
BASE_URL = "https://apiadm.study-in-egypt.gov.eg/api"
SITE_URL = "https://admission.study-in-egypt.gov.eg"
WORKER_WAIT_SECONDS = 30


def _decrypt_student_password(value):
    key = st.secrets.get("STUDENT_PASSWORD_ENCRYPTION_KEY", os.getenv("STUDENT_PASSWORD_ENCRYPTION_KEY"))
    if not key:
        raise RuntimeError("STUDENT_PASSWORD_ENCRYPTION_KEY مش موجود في Secrets.")
    return Fernet(key.encode()).decrypt(str(value).encode()).decode()


def _legacy_api_login(email, password):
    session = requests.Session()
    session.headers.update({
        "accept": "application/json, text/plain, */*",
        "accept-language": "ar",
        "device": "CITIZEN",
        "origin": SITE_URL,
        "referer": SITE_URL + "/",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36",
        "content-type": "application/json",
    })
    response = session.post(f"{BASE_URL}/student/login", json={"email": email, "password": password}, timeout=30)
    if response.status_code not in (200, 201):
        return None, None, f"فشل تسجيل الدخول - كود: {response.status_code}"
    body = response.json() if response.content else {}
    token = body.get("token", "") or response.headers.get("x-csrf-token", "")
    return session, token, None


def _legacy_api_get_status(session, token):
    filt = {
        "where": {},
        "limit": 10,
        "offset": 0,
        "order": "statusUpdatedAt DESC",
        "fields": ["serviceSlug", "ID", "createdAt", "statusUpdatedAt", "activityId", "activityName"],
    }
    headers = {"x-csrf-token": token} if token else {}
    response = session.get(
        f"{BASE_URL}/dynamic_services/inbox",
        params={"filter": json.dumps(filt)},
        headers=headers,
        timeout=30,
    )
    if response.status_code not in (200, 304):
        return f"خطأ في جلب الحالة ({response.status_code})"
    result = response.json().get("result") or []
    if not result:
        return "مفيش طلبات"
    activity = result[0].get("activityName") or "غير محدد"
    mapping = {
        "قبول الفحص الفنى": "القبول المبدئي",
        "قبول الفحص الفني": "القبول المبدئي",
        "تم السداد": "تم السداد",
        "تأكيد استلام الملف وصحة و اكتمال المستندات": "تأكيد استلام الملف وصحة واكتمال المستندات",
        "الانتظار مراجعة الطلب": "بانتظار مراجعة الطلب",
        "قبول من رئيس الادارة المركزية": "قبول من رئيس الإدارة المركزية",
    }
    return mapping.get(activity, activity)


def _legacy_api_logout(session):
    if session is None:
        return
    try:
        session.post(f"{BASE_URL}/student/logout", json={"redirectUrl": SITE_URL}, timeout=15)
    except Exception:
        pass


def _claim_fallback_job(job_id):
    rows = (
        db().table("jobs")
        .update({"status": "processing", "started_at": now_iso(), "claimed_by": "streamlit-fallback"})
        .eq("id", job_id)
        .eq("status", "pending")
        .select("*")
        .execute().data or []
    )
    return rows[0] if rows else None


def _run_legacy_api_fallback(job):
    students = (
        db().table("student_records")
        .select("*")
        .eq("office_id", job["office_id"])
        .eq("data_source_id", job["data_source_id"])
        .order("source_row_number")
        .execute().data or []
    )
    students = [
        row for row in students
        if str(row.get("application_status") or "").strip() not in FINAL_STATUSES
        and row.get("login_identifier")
        and row.get("encrypted_password")
    ]
    if not students:
        db().table("jobs").update({"status": "done", "finished_at": now_iso()}).eq("id", job["id"]).execute()
        return

    total = len(students)
    technical_failures = []

    def check_one(student):
        session = None
        status = "خطأ فني في الفحص"
        technical = False
        error_text = ""
        try:
            session, token, error = _legacy_api_login(
                student["login_identifier"],
                _decrypt_student_password(student["encrypted_password"]),
            )
            if error:
                status = "فشل تسجيل الدخول"
            else:
                status = _legacy_api_get_status(session, token)
                technical = str(status).startswith("خطأ")
                if technical:
                    status = "خطأ فني في الفحص"
        except Exception as exc:
            technical = True
            error_text = str(exc)
            status = "خطأ فني في الفحص"
        finally:
            _legacy_api_logout(session)
        return status, technical, error_text

    def save_result(student, status, index_for_progress):
        stamp = now_iso()
        student_id = student["id"]
        try:
            db().table("student_records").update({
                "application_status": status,
                "status_updated_at": stamp,
                "updated_at": stamp,
            }).eq("id", student_id).execute()
        except Exception as exc:
            print(f"Student status save error for {student_id}: {exc}")
        student_display = student.get("student_name") or student.get("login_identifier")
        try:
            db().table("job_progress").insert({
                "job_id": job["id"],
                "student_index": index_for_progress,
                "total": total,
                "student_name": student_display,
                "status": status,
            }).execute()
        except Exception as exc:
            print(f"Progress save error for {student_id}: {exc}")

    for index, student in enumerate(students, 1):
        status, technical, error_text = check_one(student)
        if technical:
            technical_failures.append({"student": student, "error": error_text or status})
        save_result(student, status, index)

    if technical_failures:
        for retry_index, item in enumerate(technical_failures, 1):
            student = item["student"]
            status, technical, error_text = check_one(student)
            save_result(student, status, total + retry_index)

    remaining_tech = []
    for item in technical_failures:
        student_id = item["student"]["id"]
        try:
            row = (
                db().table("student_records")
                .select("application_status")
                .eq("id", student_id)
                .limit(1)
                .execute().data or []
            )
            if row and str(row[0].get("application_status") or "").strip() == "خطأ فني في الفحص":
                remaining_tech.append(item)
        except Exception:
            remaining_tech.append(item)

    db().table("jobs").update({
        "status": "failed" if len(remaining_tech) >= total else "done",
        "finished_at": now_iso(),
        "error": "فشل فني في كل الطلاب" if len(remaining_tech) >= total else None,
    }).eq("id", job["id"]).execute()

def wait_for_worker_or_legacy_fallback(job_id):
    deadline=time.monotonic()+WORKER_WAIT_SECONDS
    while time.monotonic()<deadline:
        job=get_job(job_id)
        if not job: return "missing"
        if str(job.get("status") or "pending")!="pending": return "worker"
        time.sleep(2)
    claimed=_claim_fallback_job(job_id)
    if claimed:
        _run_legacy_api_fallback(claimed)
        return "fallback"
    return "worker"


def _background_update_job(job_id):
    try:
        wait_for_worker_or_legacy_fallback(job_id)
    except Exception as exc:
        try:
            db().table("jobs").update({"status":"failed","finished_at":now_iso(),"error":str(exc)[:1000]}).eq("id",job_id).execute()
        except Exception as db_exc:
            print(f"Background job save error: {db_exc}")


def _prepare_and_start_job(office_id, source_type, source_name, file_bytes, source_url, file_name):
    try:
        src, count = import_students(office_id, source_type, source_name, file_bytes=file_bytes, source_url=source_url)
        job = create_job(office_id, src, file_name)
        log_activity(office_id, "إنشاء مهمة تحديث حالات", file_name, {"job_id": job["id"], "students": count}, data_source_id=src["id"])
        t = threading.Thread(target=_background_update_job, args=(job["id"],), daemon=True)
        t.start()
        return {"job_id": job["id"], "count": count}
    except Exception as exc:
        print(f"Update preparation error: {exc}")
        return {"error": "preparation_failed"}

# ==================== UI ====================
st.set_page_config(page_title="Aivora - Agent", page_icon="✨", layout="wide", initial_sidebar_state="collapsed")

import streamlit.components.v1 as components
components.html("""
<script>(function(){try{var d=window.parent.document;if(!d.querySelector('meta[name=\"google\"][content=\"notranslate\"]')){var m=d.createElement('meta');m.name='google';m.content='notranslate';d.head.appendChild(m);}d.documentElement.classList.add('notranslate');d.documentElement.setAttribute('translate','no');}catch(e){}})();</script>
""", height=0, width=0)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800&display=swap');
html,body,[class*="css"],.stApp{font-family:'Cairo',sans-serif!important;direction:rtl;translate:no}.stApp{background:#f5f7fb;color:#111827}.block-container{max-width:1180px;padding-top:1.4rem;padding-bottom:3rem}#MainMenu,footer,header,[data-testid="stToolbar"],[data-testid="stDecoration"],[data-testid="stStatusWidget"],[data-testid="stSidebarNav"]{display:none!important}h1,h2,h3,h4,p,label,span,div{font-family:'Cairo',sans-serif!important}h1{color:#111827!important;font-size:32px!important;font-weight:800!important}.stCaption,[data-testid="stCaptionContainer"] p{color:#6b7280!important}
.topbar{background:#fff;border:1px solid #e5e7eb;border-radius:16px;padding:13px 18px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 2px 10px rgba(17,24,39,.04);margin-bottom:22px}.brand{display:flex;align-items:center;gap:11px}.brand-icon{width:42px;height:42px;border-radius:12px;background:#eff6ff;display:flex;align-items:center;justify-content:center;font-size:22px}.brand-title{font-size:18px;font-weight:800;color:#111827}.brand-sub{font-size:12px;color:#6b7280}.card{background:#fff;border:1px solid #e5e7eb;border-radius:18px;padding:22px;box-shadow:0 3px 14px rgba(17,24,39,.045);margin-bottom:18px}.hero{background:linear-gradient(135deg,#fff 0%,#f8fbff 100%);border:1px solid #dbeafe;border-radius:20px;padding:25px 28px;box-shadow:0 4px 18px rgba(37,99,235,.06);margin-bottom:20px}.hero-kicker{color:#6b7280;font-size:14px;font-weight:600}.hero-title{color:#111827;font-size:28px;font-weight:800;margin-top:2px}.hero-title strong{color:#2563eb}.hero-desc{color:#6b7280;font-size:14px;margin-top:3px}.stat-card{background:#fff;border:1px solid #e5e7eb;border-radius:16px;padding:17px 18px}.stat-label{color:#6b7280;font-size:12px;font-weight:600}.stat-value{color:#111827;font-size:21px;font-weight:800;margin-top:2px}.section-title{font-size:18px;font-weight:800;color:#111827;margin:5px 0 13px}.section-sub{color:#6b7280;font-size:13px;margin-top:-8px;margin-bottom:14px}
.stTextInput label,.stFileUploader label,.stRadio>label,.stCheckbox label{color:#374151!important;font-size:14px!important;font-weight:700!important}.stTextInput input{background:#fff!important;color:#111827!important;border:1px solid #d1d5db!important;border-radius:10px!important;font-size:14px!important;min-height:44px}.stTextInput input:focus{border-color:#2563eb!important;box-shadow:0 0 0 3px rgba(37,99,235,.10)!important}.stTextInput input::placeholder{color:#9ca3af!important}.stButton>button,.stDownloadButton>button{width:100%;min-height:44px;border-radius:10px!important;border:1px solid #2563eb!important;background:#2563eb!important;color:#fff!important;font-weight:700!important;font-size:14px!important;box-shadow:0 3px 8px rgba(37,99,235,.16)!important;transition:.15s ease}.stButton>button:hover,.stDownloadButton>button:hover{background:#1d4ed8!important;border-color:#1d4ed8!important;transform:translateY(-1px)}button[kind="secondary"]{background:#fff!important;color:#2563eb!important}.stRadio div[role="radiogroup"]{gap:10px}.stRadio div[role="radiogroup"] label{background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:10px 14px}[data-testid="stFileUploaderDropzone"]{background:#f8fafc!important;border:1.5px dashed #cbd5e1!important;border-radius:14px!important}[data-testid="stFileUploaderDropzone"] button{background:#fff!important;border:1px solid #bfdbfe!important;border-radius:8px!important;font-size:0!important;line-height:1!important;white-space:nowrap!important;overflow:hidden!important}[data-testid="stFileUploaderDropzone"] button span{font-size:0!important}[data-testid="stFileUploaderDropzone"] button::after{content:"";font-family:'Cairo',sans-serif!important;font-size:13px!important;font-weight:700!important;line-height:1.2!important;color:#2563eb!important}.result-card{background:#fff;border:1px solid #e5e7eb;border-radius:13px;padding:14px 16px;margin:8px 0}.result-name{color:#111827;font-size:15px;font-weight:800}.result-status{color:#2563eb;font-size:13px;font-weight:700;margin-top:2px}.status-badge{display:inline-block;padding:4px 10px;border-radius:999px;font-size:12px;font-weight:700}.status-ok{background:#ecfdf5;color:#15803d}.status-warn{background:#fffbeb;color:#b45309}.status-error{background:#fef2f2;color:#b91c1c}.status-info{background:#eff6ff;color:#1d4ed8}.connected-box{background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;padding:11px 14px;color:#166534;font-size:13px}.success-box{background:#ecfdf5;border:1px solid #bbf7d0;border-radius:14px;padding:16px;margin-top:15px}.success-title{font-size:17px;font-weight:800;color:#166534}.success-desc{color:#166534;font-size:13px;margin-top:4px}.job-box{background:#eff6ff;border:1px solid #bfdbfe;border-radius:10px;padding:10px 14px;color:#1d4ed8;font-size:13px;margin-bottom:10px}[data-testid="InputInstructions"]{display:none!important}[data-testid="stTextInput"] button{display:none!important}
.progress-table{width:100%;border-collapse:collapse;margin-top:12px;direction:rtl}.progress-table th,.progress-table td{border-bottom:1px solid #e5e7eb;padding:10px 12px;text-align:right;font-size:13px}.progress-table th{background:#f8fafc;color:#374151;font-weight:800}.progress-table td:first-child{text-align:left;width:60px;color:#6b7280}.progress-table td:last-child{width:38%;font-weight:700}.progress-table .status-cell{text-align:left;font-weight:700}.progress-table .status-name{text-align:right;font-weight:800;color:#111827}
@media(max-width:700px){.block-container{padding:.8rem .7rem 2rem}.hero-title{font-size:23px}.topbar{padding:11px 13px}}
</style>
""", unsafe_allow_html=True)

for key,default in [("logged_in",False),("is_admin",False),("office",None),("update_locked",False),("pending_file_bytes",None),("pending_filename","") ,("active_job_id",None),("update_waiting",False),("update_starting",False),("update_error",False)]:
    if key not in st.session_state: st.session_state[key]=default

# ==================== Login (legacy layout) ====================
if not st.session_state.logged_in and not st.session_state.is_admin:
    st.markdown("<div style='text-align:center;margin:42px 0 24px;'><div style='font-size:48px;'>✨</div><div style='font-size:31px;font-weight:800;color:#111827;'>Aivora</div><div style='font-size:14px;color:#6b7280;margin-top:3px;'>Your Smarter Support for Every Student's Application</div></div>",unsafe_allow_html=True)
    left,right=st.columns([1.15,1],gap="large")
    with left:
        st.markdown("<div style='padding:40px 20px 20px 10px;'><div style='font-size:14px;color:#2563eb;font-weight:800;margin-bottom:8px;'>حل بسيط لإدارة المتابعة</div><div style='font-size:31px;font-weight:800;color:#111827;line-height:1.35;'>تابع طلبات طلابك<br>من مكان واحد.</div><div style='font-size:15px;color:#6b7280;line-height:1.9;margin-top:12px;max-width:480px;'>حدّث حالات الطلبات، اربط Google Sheets، وابحث عن أي طالب بسرعة بدون متابعة يدوية.</div><div style='margin-top:22px;color:#374151;font-size:14px;line-height:2.2;'>✓ تحديث حالات الطلاب بشكل منظم<br>✓ حفظ مصدر البيانات للمكتب<br>✓ بحث سريع عن حالة أي طالب</div></div>",unsafe_allow_html=True)
    with right:
        st.markdown('<div class="card">',unsafe_allow_html=True)
        tab1,tab2=st.tabs(["تسجيل الدخول","حساب جديد"])
        with tab1:
            st.markdown("<div style='font-size:22px;font-weight:800;color:#111827;margin:8px 0 3px;'>مرحبًا بعودتك 👋</div><div style='color:#6b7280;font-size:13px;margin-bottom:18px;'>سجل دخولك لإدارة طلبات الطلاب</div>",unsafe_allow_html=True)
            username=st.text_input("اسم المكتب",key="login_user",placeholder="اكتب اسم المكتب")
            password=st.text_input("كلمة المرور",type="password",key="login_pass",placeholder="اكتب كلمة المرور")
            if st.button("تسجيل الدخول",key="login_btn"):
                if username==ADMIN_USERNAME and password==ADMIN_PASSWORD:
                    st.session_state.is_admin=True; st.rerun()
                else:
                    ok,result=check_login(username,password)
                    if ok: st.session_state.logged_in=True; st.session_state.office=result; st.session_state.update_locked=False; st.session_state.update_waiting=False; st.session_state.update_starting=False; st.session_state.update_error=False; st.rerun()
                    else: st.error(result)
        with tab2:
            st.markdown("<div style='font-size:22px;font-weight:800;color:#111827;margin:8px 0 3px;'>إنشاء حساب</div><div style='color:#6b7280;font-size:13px;margin-bottom:18px;'>سجل مكتبك لبدء استخدام الخدمة</div>",unsafe_allow_html=True)
            new_office=st.text_input("اسم المكتب",key="reg_office",placeholder="اسم المكتب")
            new_email=st.text_input("الإيميل",key="reg_email",placeholder="example@email.com")
            new_pass=st.text_input("كلمة المرور",type="password",key="reg_pass",placeholder="كلمة المرور")
            new_pass2=st.text_input("تأكيد كلمة المرور",type="password",key="reg_pass2",placeholder="أعد كتابة كلمة المرور")
            if st.button("إنشاء الحساب",key="reg_btn"):
                if not new_office or not new_email or not new_pass: st.error("املأ كل الحقول!")
                elif new_pass!=new_pass2: st.error("كلمة المرور مش متطابقة!")
                else:
                    ok,msg=register_office(new_office,new_email,new_pass); (st.success if ok else st.error)(msg)
        st.markdown('</div>',unsafe_allow_html=True)
    st.stop()

# ==================== Admin ====================
if st.session_state.is_admin:
    st.markdown("<div class='hero'><div class='hero-kicker'>الإدارة</div><div class='hero-title'>لوحة الإدارة</div><div class='hero-desc'>إدارة الحسابات الجديدة ومتابعة حالة المكاتب.</div></div>",unsafe_allow_html=True)
    pending=get_pending_accounts(); st.markdown(f"<div class='section-title'>الحسابات المعلقة <span style='color:#2563eb'>({len(pending)})</span></div>",unsafe_allow_html=True)
    if pending:
        for acc in pending:
            c1,c2,c3,c4=st.columns([3,3,1,1]); c1.write(acc.get("name","")); c2.write(acc.get("email",""))
            if c3.button("قبول",key=f"approve_{acc['id']}"): set_office_status(acc["id"],"approved"); log_activity(acc["id"],"تمت الموافقة على حساب المكتب"); st.rerun()
            if c4.button("رفض",key=f"reject_{acc['id']}"): set_office_status(acc["id"],"rejected"); log_activity(acc["id"],"تم رفض حساب المكتب"); st.rerun()
    else: st.info("لا توجد حسابات في انتظار الموافقة.")
    st.markdown("<div class='section-title' style='margin-top:28px;'>كل الحسابات</div>",unsafe_allow_html=True)
    try:
        all_accounts=db().table("offices").select("name,email,status,created_at").order("created_at").execute().data or []
        if all_accounts: st.dataframe(pd.DataFrame(all_accounts),use_container_width=True,hide_index=True)
    except Exception: pass
    if st.button("تسجيل الخروج من الإدارة",key="admin_logout"): st.session_state.clear(); st.rerun()
    st.stop()

# ==================== Office page ====================
office=st.session_state.office
if not office: st.session_state.logged_in=False; st.rerun()
office_id=office["id"]

st.markdown("<div class='topbar'><div class='brand'><div class='brand-icon'>✨</div><div><div class='brand-title'>Aivora</div><div class='brand-sub'>Your Smarter Support for Every Student's Application</div></div></div></div>", unsafe_allow_html=True)

# ==================== Data source card ====================
st.markdown('<div class="card">',unsafe_allow_html=True)
st.markdown('<div class="section-title">مصدر بيانات الطلاب</div><div class="section-sub">اختر الطريقة التي يحتوي بها ملف الطلاب.</div>',unsafe_allow_html=True)
source_options=["📂 رفع ملف Excel","🔗 ربط Google Sheets"]
source=st.radio("",source_options,horizontal=True,label_visibility="collapsed",key="source_mode")
file_bytes=None; filename=""; saved_link=get_saved_gsheet_link(office_id); sheet_id_source=None

if source=="📂 رفع ملف Excel":
    uploaded = st.file_uploader(
        "ارفع ملف Excel",
        type=["xlsx", "xls","csv"],
        label_visibility="collapsed",
        key="excel_upload",
    )
    if uploaded:
        file_bytes = uploaded.getvalue()
        filename = uploaded.name
        if uploaded.name.lower().endswith(".csv"):
            df = pd.read_csv(io.BytesIO(file_bytes), encoding="utf-8-sig")
        else:
            df = pd.read_excel(io.BytesIO(file_bytes))
        st.success(f"تم اختيار الملف: {uploaded.name}")
else:
    if saved_link:
        st.markdown('<div class="connected-box">✓ Google Sheets متصل بالفعل لهذا المكتب</div>',unsafe_allow_html=True); st.markdown("<div style='height:8px'></div>",unsafe_allow_html=True)
        change=st.checkbox("تغيير رابط الشيت",key="change_link")
        if change:
            new_link=st.text_input("رابط Google Sheets",placeholder="الصق الرابط هنا",key="new_link")
            if st.button("حفظ الرابط",key="save_link"):
                sid=extract_sheet_id(new_link) if new_link else None
                if not new_link: st.error("أدخل الرابط أولاً!")
                elif not sid: st.error("الرابط غير صحيح!")
                else:
                    ok,msg=save_gsheet_link(office_id,new_link); st.success("تم حفظ الرابط!") if ok else st.error(msg)
                    if ok: st.rerun()
        sheet_id_source=extract_sheet_id(saved_link)
    else:
        new_link=st.text_input("رابط Google Sheets",placeholder="https://docs.google.com/spreadsheets/d/...",key="first_link")
        if st.button("حفظ وربط الشيت",key="save_first_link"):
            sid=extract_sheet_id(new_link) if new_link else None
            if sid:
                ok,msg=save_gsheet_link(office_id,new_link); st.success("تم حفظ الرابط بنجاح!") if ok else st.error(msg)
                if ok: st.rerun()
            else: st.error("الرابط غير صحيح!")
    if sheet_id_source:
        update_option=st.radio("طريقة التحديث",["🔄 تحديث من الشيت الأونلاين","📂 رفع Excel جديد بدلاً عنه"],horizontal=True,key="update_option")
        if update_option=="🔄 تحديث من الشيت الأونلاين":
            if st.button("تحميل بيانات الشيت",key="load_sheet"):
                with st.spinner("بيقرأ الشيت..."):
                    rows=read_gsheet_rows(saved_link)
                    if not rows: st.error("الشيت فاضي!")
                    else:
                        wb=openpyxl.Workbook(); ws=wb.active
                        for r in rows: ws.append(r)
                        out=io.BytesIO(); wb.save(out)
                        st.session_state.pending_file_bytes=out.getvalue(); st.session_state.pending_filename="google_sheet"; st.success("تم جلب بيانات الشيت. البيانات جاهزة للتحديث.")
        else:
            uploaded2=st.file_uploader("ارفع ملف Excel الجديد",type=["xlsx","xls"],label_visibility="collapsed",key="excel_replace")
            if uploaded2: file_bytes=uploaded2.getvalue(); filename=uploaded2.name; st.session_state.pending_file_bytes=file_bytes; st.session_state.pending_filename=filename

if not file_bytes and st.session_state.pending_file_bytes: file_bytes=st.session_state.pending_file_bytes; filename=st.session_state.pending_filename or filename
st.markdown('</div>',unsafe_allow_html=True)

# ==================== Processing ====================
@st.fragment(run_every=2)
def render_processing():
    if not file_bytes and not st.session_state.active_job_id and not st.session_state.update_waiting: return
    st.markdown('<div class="card">',unsafe_allow_html=True)
    st.markdown('<div class="section-title">تحديث حالات الطلاب</div><div class="section-sub">اضغط الزر لبدء فحص الطلبات وتحديث النتائج.</div>',unsafe_allow_html=True)
    if st.session_state.update_waiting and not st.session_state.active_job_id:
        st.info("⏳ سيبدأ التحديث خلال ثواني…")
    elif st.session_state.update_locked and not st.session_state.active_job_id:
        st.info("🔒 تم تشغيل تحديث بالفعل في هذه الجلسة. لو عايزة تبدئي تحديث جديد، سجّلي خروج وادخلي تاني.")
    elif file_bytes and not st.session_state.update_locked and not st.session_state.active_job_id:
        if st.button("▶ تحديث حالات الطلاب",key="start_main"):
            running=db().table("jobs").select("id").eq("office_id",office_id).in_("status",["pending","processing"]).limit(1).execute().data or []
            if running:
                st.warning("في تحديث شغال بالفعل لهذا المكتب. استني لحد ما يخلص.")
            else:
                st.session_state.update_locked=True
                st.session_state.update_waiting=True
                st.session_state.update_error=False
                is_gsheet_source=bool(saved_link and source=="🔗 ربط Google Sheets")
                source_type="gsheet" if is_gsheet_source else "xlsx"
                source_name="Google Sheet" if is_gsheet_source else (filename or "students.xlsx")
                prep=threading.Thread(target=_prepare_and_start_job,args=(office_id,source_type,source_name,file_bytes,saved_link if is_gsheet_source else None,filename or source_name),daemon=True)
                prep.start()
                st.session_state.update_starting=prep
    job=get_job(st.session_state.active_job_id) if st.session_state.active_job_id else None
    if not job and st.session_state.update_waiting:
        recent=db().table("jobs").select("*").eq("office_id",office_id).order("created_at",desc=True).limit(1).execute().data or []
        if recent:
            candidate=recent[0]
            if str(candidate.get("status") or "pending") in {"pending","processing","done","failed"}:
                st.session_state.active_job_id=candidate["id"]
                st.session_state.update_waiting=False
                job=candidate
    if job:
        status=str(job.get("status") or "pending")
        if status=="pending":
            st.info("⏳ سيبدأ التحديث خلال ثواني…")
        elif status=="processing":
            st.success("▶️ بدأ التحديث. جاري فحص الطلاب وتحديث الحالات…")
        rows=get_job_progress_rows(job["id"])
        if rows:
            latest={}
            for r in rows:
                key=(str(r.get("student_name") or "").strip().lower(),str(r.get("status") or "").strip())
                if key[0]: latest[key]=r
            shown=sorted(latest.values(),key=lambda r:int(r.get("student_index") or 0),reverse=True)
            total=max(int(r.get("total") or 0) for r in rows)
            completed=min(len({str(r.get("student_name") or "").strip().lower() for r in rows if str(r.get("student_name") or "").strip()}),total) if total else 0
            st.progress(min(completed/max(total,1),1.0))
            st.caption(f"طالب {completed} من {total}")
            last=sorted(rows,key=lambda r:int(r.get("student_index") or 0))[-1]
            st.info(f"🔄 آخر طالب تم فحصه: **{last.get('student_name') or 'طالب'}** — الحالة الجديدة: **{last.get('status') or ''}**")
            table_rows=[]
            seen_names=set()
            for r in sorted(rows,key=lambda x:int(x.get("student_index") or 0),reverse=True):
                name=str(r.get("student_name") or "").strip()
                norm=name.lower()
                if not name or norm in seen_names: continue
                seen_names.add(norm)
                table_rows.append(f"<tr><td class='status-cell'>{r.get('status') or ''}</td><td class='status-name'>{name}</td><td>{r.get('student_index') or ''}</td></tr>")
            st.markdown(f"<table class='progress-table'><thead><tr><th>الحالة</th><th>اسم الطالب</th><th>الرقم</th></tr></thead><tbody>{''.join(table_rows)}</tbody></table>",unsafe_allow_html=True)
        if status=="done":
            st.markdown('<div class="success-box"><div class="success-title">اكتمل التحديث 🎉</div><div class="success-desc">تمت معالجة الطلبات وتحديث الحالات.</div></div>',unsafe_allow_html=True)
        elif status=="failed":
            st.error("حدثت مشكلة أثناء تحديث الطلبات. بعض الحالات قد تحتاج إلى إعادة المحاولة.")
    st.markdown('</div>',unsafe_allow_html=True)

render_processing()

# ==================== Search ====================
st.markdown('<div class="card">',unsafe_allow_html=True)
st.markdown('<div class="section-title">البحث عن طالب</div><div class="section-sub">اكتب اسم الطالب لمعرفة آخر حالة محفوظة.</div>',unsafe_allow_html=True)
search_query=st.text_input("اسم الطالب",placeholder="مثال: Ahmed Mohamed",label_visibility="collapsed",key="student_search")
if search_query:
    found=get_students(office_id,search_query)
    if found:
        for s in found:
            status=s.get("application_status") or "لم يتم الفحص بعد"
            st.markdown(f'<div class="result-card"><div class="result-name">👤 {s.get("student_name","")}</div><div class="result-status"><span class="status-badge {status_class(status)}">{status}</span></div></div>',unsafe_allow_html=True)
    else: st.info("مفيش طالب بالاسم ده.")
st.markdown('</div>',unsafe_allow_html=True)

if st.button("تسجيل الخروج",key="logout_main"): st.session_state.clear(); st.rerun()

log_activity(office_id,"فتح لوحة المكتب")
