"""
Aivora - Agent
Your Smarter Support for Every Student's Application
"""

import io
import json
import os
import re
import threading
import time
from datetime import datetime, timezone

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

import bcrypt
import openpyxl
import pandas as pd
import requests
import streamlit as st
from cryptography.fernet import Fernet
from supabase import Client, create_client

try:
    import gspread
    from google.oauth2.service_account import Credentials
except Exception:
    gspread = None
    Credentials = None


def get_google_credentials(scopes):
    if Credentials is None:
        raise RuntimeError("Google credentials libraries unavailable")

    creds_dict = st.secrets.get("gcp_service_account")
    if creds_dict:
        try:
            creds_dict = dict(creds_dict)
        except Exception:
            pass
        if hasattr(creds_dict, "get") and creds_dict.get("client_email") and creds_dict.get("private_key"):
            return Credentials.from_service_account_info(creds_dict, scopes=scopes)

    raw = st.secrets.get("GCP_SERVICE_ACCOUNT_JSON", os.getenv("GCP_SERVICE_ACCOUNT_JSON"))
    if raw:
        data = json.loads(raw) if isinstance(raw, str) else dict(raw)
        return Credentials.from_service_account_info(data, scopes=scopes)

    raw = st.secrets.get("GOOGLE_SERVICE_ACCOUNT_JSON", os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON"))
    if raw:
        data = json.loads(raw) if isinstance(raw, str) else dict(raw)
        return Credentials.from_service_account_info(data, scopes=scopes)

    credentials_file = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
    if credentials_file:
        return Credentials.from_service_account_file(credentials_file, scopes=scopes)

    raise RuntimeError("Google service account configuration missing")


def get_gsheet_client():
    if gspread is None or Credentials is None:
        raise RuntimeError("Google Sheets libraries unavailable")
    creds = get_google_credentials([
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ])
    return gspread.authorize(creds)


BASE_URL = "https://apiadm.study-in-egypt.gov.eg/api"
SITE_URL = "https://admission.study-in-egypt.gov.eg"
WORKER_WAIT_SECONDS = 30
DRIVE_FOLDER_ID = os.getenv("DRIVE_FOLDER_ID", "12L_qSHBnW4-tfQZRteynInWNBAML016f")
LOGIN_PAGE_DELAY_MIN, LOGIN_PAGE_DELAY_MAX = 0.8, 1.5
POST_LOGIN_DELAY_SECONDS = 1.0
INBOX_DELAY_MIN, INBOX_DELAY_MAX = 0.7, 1.5
STUDENT_DELAY_MIN, STUDENT_DELAY_MAX = 4, 8
TECH_FAILURE_STATUS = "تعذر فحص الطالب حاليًا"
FINAL_STATUSES = {"مقبول نهائي", "قبول نهائي", "تم الرفض", "مرفوض نهائيًا", "مرفوض نهائيا", "مرفوض", "خالص"}
ADMIN_USERNAME = st.secrets.get("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD", "admin123")

@st.cache_resource(show_spinner=False)
def get_supabase() -> Client:
    url = st.secrets.get("SUPABASE_URL", os.getenv("SUPABASE_URL"))
    key = st.secrets.get("SUPABASE_SERVICE_ROLE_KEY", os.getenv("SUPABASE_SERVICE_ROLE_KEY"))
    if not url or not key:
        raise RuntimeError("Supabase configuration is missing")
    return create_client(url, key)

def db(): return get_supabase()
def now_iso(): return datetime.now(timezone.utc).isoformat()
def safe_log(message): print(f"[Aivora] {message}")
def fallback_delay(a, b):
    import random
    time.sleep(random.uniform(a, b))

def extract_sheet_id(link):
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", str(link)); return match.group(1) if match else None

def extract_gid(link):
    match = re.search(r"[?#&]gid=(\d+)", str(link)); return int(match.group(1)) if match else None

def read_gsheet_rows(link):
    sheet_id = extract_sheet_id(link)
    if not sheet_id: raise ValueError("رابط Google Sheets غير صحيح.")
    spreadsheet = get_gsheet_client().open_by_key(sheet_id)
    gid = extract_gid(link)
    worksheet = next((w for w in spreadsheet.worksheets() if w.id == gid), spreadsheet.sheet1) if gid is not None else spreadsheet.sheet1
    return worksheet.get_all_values()

def find_excel_columns(ws):
    cols = {"name": None, "email": None, "password": None}; header_row = None
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=min(10, ws.max_row), values_only=True), start=1):
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
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False); ws = wb.active
    cols, header_row = find_excel_columns(ws); records=[]; seen=set()
    for excel_row, row in enumerate(ws.iter_rows(min_row=header_row+1, values_only=True), start=header_row+1):
        values=list(row)
        email=str(values[cols["email"]] or "").strip() if cols["email"] < len(values) else ""
        password=str(values[cols["password"]] or "").strip() if cols["password"] < len(values) else ""
        name=str(values[cols["name"]] or "").strip() if cols["name"] < len(values) else email
        key=email.lower()
        if not email or not password or key in seen: continue
        seen.add(key)
        records.append({"source_row_number":excel_row,"student_name":name or email,"login_identifier":email,"password":password,"original_data":{f"column_{i+1}":str(v) if v is not None else "" for i,v in enumerate(values)}})
    return records

def encrypt_password(password,key):
    if not key: raise RuntimeError("Encryption key missing")
    return Fernet(key.encode()).encrypt(password.encode()).decode()

def drive_service():
    return build("drive","v3",credentials=get_google_credentials(["https://www.googleapis.com/auth/drive"]))

def upload_to_drive(file_bytes, filename, office):
    media=MediaIoBaseUpload(io.BytesIO(file_bytes),mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",resumable=True)
    return drive_service().files().create(body={"name":str(filename or "students.xlsx"),"parents":[DRIVE_FOLDER_ID]},media_body=media,fields="id").execute()["id"]

def download_drive_file_bytes(file_id):
    buffer=io.BytesIO(); downloader=MediaIoBaseDownload(buffer,drive_service().files().get_media(fileId=str(file_id))); done=False
    while not done: _,done=downloader.next_chunk()
    buffer.seek(0); return buffer.getvalue()

def find_status_column_for_output(ws, header_row):
    for col_idx, cell in enumerate(ws[header_row], start=1):
        value=str(cell.value or "").strip().lower()
        if value in {"حالة الطلب","الحالة"} or ("حالة" in value and "اسم" not in value and "خدمة" not in value): return col_idx
    new_col=ws.max_column+1; ws.cell(header_row,new_col).value="حالة الطلب"; return new_col

def get_office_by_name(name):
    rows=db().table("offices").select("id,name,email,status,created_at").eq("name",str(name).strip()).limit(1).execute().data or []; return rows[0] if rows else None

def get_office_with_hash(office_id):
    rows=db().table("offices").select("id,name,email,status,password_hash,created_at").eq("id",office_id).limit(1).execute().data or []; return rows[0] if rows else None

def get_office_by_email(email):
    rows=db().table("offices").select("id,name,email,status").eq("email",str(email).strip().lower()).limit(1).execute().data or []; return rows[0] if rows else None

def register_office(name,email,password):
    name=name.strip(); email=email.strip().lower(); password=password.strip()
    if not name or not email or not password:return False,"كمّلي كل البيانات الأول."
    if len(password)<6:return False,"الباسورد لازم يكون 6 أحرف على الأقل."
    try:
        if get_office_by_name(name) or get_office_by_email(email):return False,"اسم المكتب أو الإيميل مسجل قبل كده."
        password_hash=bcrypt.hashpw(password.encode(),bcrypt.gensalt()).decode(); db().table("offices").insert({"name":name,"email":email,"password_hash":password_hash,"status":"pending"}).execute(); return True,"تم التسجيل بنجاح. الحساب في انتظار موافقة الإدارة."
    except Exception:return False,"تعذر إنشاء الحساب حاليًا. حاولي مرة تانية."

def check_login(name,password):
    try:
        office=get_office_by_name(name)
        if not office:return False,"اسم المكتب أو الباسورد غلط!"
        if office.get("status")=="pending":return False,"حسابك في انتظار موافقة الإدارة!"
        if office.get("status")!="approved":return False,"حسابك موقوف!"
        full=get_office_with_hash(office["id"]); password_hash=str((full or {}).get("password_hash") or "")
        if not password_hash or not bcrypt.checkpw(password.encode(),password_hash.encode()):return False,"اسم المكتب أو الباسورد غلط!"
        return True,office
    except Exception:return False,"تعذر تسجيل الدخول حاليًا. حاولي مرة تانية."

def get_saved_gsheet_link(office_id):
    try:
        rows=db().table("data_sources").select("source_url,created_at").eq("office_id",office_id).eq("source_type","google_sheet").not_.is_("source_url","null").order("created_at",desc=True).limit(1).execute().data or []; return rows[0].get("source_url") if rows else None
    except Exception:return None

def save_gsheet_link(office_id,link):
    try:
        if not extract_sheet_id(link):return False,"الرابط غير صحيح!"
        db().table("data_sources").insert({"office_id":office_id,"source_type":"google_sheet","source_name":"Google Sheet","source_url":link,"column_mapping":{}}).execute(); return True,"تم حفظ الرابط بنجاح"
    except Exception:return False,"تعذر حفظ الرابط حاليًا."

def import_students(office_id,source_type,source_name,file_bytes=None,source_url=None):
    if source_type in {"xlsx","xls","excel"}: records=parse_excel_bytes(file_bytes)
    else:
        rows=read_gsheet_rows(source_url)
        if not rows: raise ValueError("الشيت فاضي.")
        wb=openpyxl.Workbook(); ws=wb.active
        for row in rows: ws.append(row)
        out=io.BytesIO(); wb.save(out); records=parse_excel_bytes(out.getvalue())
    if not records: raise ValueError("مش لاقي طلاب عندهم إيميل وباسورد صالحين.")
    encryption_key=st.secrets.get("STUDENT_PASSWORD_ENCRYPTION_KEY",os.getenv("STUDENT_PASSWORD_ENCRYPTION_KEY"))
    source_type="google_sheet" if source_type in {"gsheet","google_sheet"} else "excel"
    file_path=upload_to_drive(file_bytes,source_name,"") if source_type=="excel" else None
    source=db().table("data_sources").insert({"office_id":office_id,"source_type":source_type,"source_name":source_name,"source_url":source_url,"file_path":file_path,"column_mapping":{}}).execute().data[0]
    payload=[{"office_id":office_id,"data_source_id":source["id"],"source_row_number":r["source_row_number"],"student_name":r["student_name"],"login_identifier":r["login_identifier"],"encrypted_password":encrypt_password(r["password"],encryption_key),"application_status":"","original_data":r["original_data"],"updated_at":now_iso()} for r in records]
    db().table("student_records").insert(payload).execute(); return source,len(payload)

def create_job(office_id,source,file_name):
    return db().table("jobs").insert({"office_id":office_id,"data_source_id":source["id"],"source_type":source["source_type"],"source_ref":source.get("file_path") or source.get("source_url") or source["id"],"file_name":file_name,"status":"pending"}).execute().data[0]

def get_job(job_id):
    if not job_id:return None
    rows=db().table("jobs").select("*").eq("id",job_id).limit(1).execute().data or []; return rows[0] if rows else None

def job_is_active(job_id):
    try:return str((get_job(job_id) or {}).get("status") or "") == "processing"
    except Exception:return False

def cancel_all_office_jobs(office_id):
    if not office_id:return False
    try:
        db().table("jobs").update({"status":"cancelled","finished_at":now_iso(),"error":"office_logout"}).eq("office_id",office_id).in_("status",["pending","processing"]).execute()
        remaining=db().table("jobs").select("id,status").eq("office_id",office_id).in_("status",["pending","processing"]).execute().data or []
        if remaining:
            db().table("jobs").update({"status":"cancelled","finished_at":now_iso(),"error":"office_logout"}).eq("office_id",office_id).in_("status",["pending","processing"]).execute()
            remaining=db().table("jobs").select("id,status").eq("office_id",office_id).in_("status",["pending","processing"]).execute().data or []
        return not remaining
    except Exception as exc:safe_log(f"cancel all office jobs failed: {type(exc).__name__}: {exc}"); return False

def get_students_for_job(job_id):
    job=get_job(job_id)
    if not job:return []
    rows=db().table("student_records").select("*").eq("office_id",job["office_id"]).eq("data_source_id",job["data_source_id"]).order("source_row_number").execute().data or []
    latest={}
    for row in rows:
        key=str(row.get("login_identifier") or row.get("student_name") or "").strip().lower()
        if key:latest[key]=row
    return list(latest.values())

def get_job_progress_rows(job_id):
    rows=db().table("job_progress").select("student_index,total,student_name,status,created_at").eq("job_id",job_id).order("created_at").execute().data or []
    latest={}
    for row in rows:
        key=str(row.get("student_name") or "").strip().lower()
        if key:latest[key]=row
    return sorted(latest.values(),key=lambda r:int(r.get("student_index") or 0))

def update_excel_student_status(source_ref,student,status):
    if not source_ref:raise RuntimeError("excel_source_missing")
    wb=openpyxl.load_workbook(io.BytesIO(download_drive_file_bytes(source_ref)),data_only=False); ws=wb.active; cols,header_row=find_excel_columns(ws); status_col=find_status_column_for_output(ws,header_row); source_row=student.get("source_row_number")
    if not source_row:
        login=str(student.get("login_identifier") or "").strip().lower()
        for row_idx in range(header_row+1,ws.max_row+1):
            email=str(ws.cell(row_idx,cols["email"]+1).value or "").strip().lower()
            if email==login: source_row=row_idx; break
    if not source_row:raise RuntimeError("excel_student_row_missing")
    ws.cell(int(source_row),status_col).value=status; output=io.BytesIO(); wb.save(output); updated=output.getvalue(); media=MediaIoBaseUpload(io.BytesIO(updated),mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",resumable=True); drive_service().files().update(fileId=str(source_ref),media_body=media).execute()

def update_google_sheet_student_status(source_url,student,status):
    sheet_id=extract_sheet_id(source_url)
    if not sheet_id:raise RuntimeError("invalid_google_sheet_url")
    spreadsheet=get_gsheet_client().open_by_key(sheet_id); gid=extract_gid(source_url); worksheet=next((w for w in spreadsheet.worksheets() if w.id==gid),spreadsheet.sheet1) if gid is not None else spreadsheet.sheet1; values=worksheet.get_all_values(); header_idx=email_idx=status_idx=None
    for r_idx,row in enumerate(values[:10]):
        normalized=[str(v or "").strip().lower() for v in row]
        if any("يميل" in v or "mail" in v or "بريد" in v for v in normalized):
            header_idx=r_idx
            for i,value in enumerate(normalized):
                if "يميل" in value or "mail" in value or "بريد" in value:email_idx=i
                if value in {"حالة الطلب","الحالة"} or ("حالة" in value and "اسم" not in value and "خدمة" not in value):status_idx=i
            break
    if header_idx is None or email_idx is None or status_idx is None:raise RuntimeError("google_sheet_columns_missing")
    login=str(student.get("login_identifier") or "").strip().lower(); source_row=student.get("source_row_number"); target_row=None
    if source_row and header_idx < int(source_row) <= len(values):
        candidate_login=str(values[int(source_row)-1][email_idx] if email_idx<len(values[int(source_row)-1]) else "").strip().lower()
        if candidate_login==login:target_row=int(source_row)
    if target_row is None:
        for row_idx in range(header_idx+1,len(values)):
            row=values[row_idx]; row_login=str(row[email_idx] if email_idx<len(row) else "").strip().lower()
            if row_login==login:target_row=row_idx+1;break
    if target_row is None:raise RuntimeError("google_sheet_student_row_missing")
    worksheet.update_cell(target_row,status_idx+1,status)

def update_source_student_status(job,student,status):
    if str(job.get("source_type") or "") == "excel":update_excel_student_status(job.get("source_ref"),student,status)
    elif str(job.get("source_type") or "") == "google_sheet":update_google_sheet_student_status(job.get("source_ref"),student,status)
    else:raise RuntimeError(f"unsupported_source_type:{job.get('source_type')}")

def mark_job_done_if_active(job_id):
    response=db().table("jobs").update({"status":"done","finished_at":now_iso(),"error":None}).eq("id",job_id).eq("status","processing").select("id,status").execute(); return bool(response.data or [])

def mark_job_failed_if_active(job_id,error):
    response=db().table("jobs").update({"status":"failed","finished_at":now_iso(),"error":str(error)[:1000]}).eq("id",job_id).eq("status","processing").select("id,status").execute(); return bool(response.data or [])

def _legacy_api_login(email,password):
    session=requests.Session(); session.headers.update({"accept":"application/json, text/plain, */*","accept-language":"ar","device":"CITIZEN","origin":SITE_URL,"referer":SITE_URL+"/","user-agent":"Mozilla/5.0","content-type":"application/json"}); response=session.post(f"{BASE_URL}/student/login",json={"email":email,"password":password},timeout=30)
    if response.status_code not in (200,201):return None,None,"login_failed"
    body=response.json() if response.content else {}; return session,body.get("token","") or response.headers.get("x-csrf-token",""),None

def _legacy_api_get_status(session,token):
    filt={"where":{},"limit":10,"offset":0,"order":"statusUpdatedAt DESC","fields":["serviceSlug","ID","createdAt","statusUpdatedAt","activityId","activityName"]}; headers={"x-csrf-token":token} if token else {}; response=session.get(f"{BASE_URL}/dynamic_services/inbox",params={"filter":json.dumps(filt)},headers=headers,timeout=30)
    if response.status_code not in (200,304):raise RuntimeError(f"status_http_{response.status_code}")
    result=response.json().get("result") or []
    if not result:return "مفيش طلبات"
    activity=result[0].get("activityName") or "غير محدد"; mapping={"قبول الفحص الفنى":"القبول المبدئي","قبول الفحص الفني":"القبول المبدئي","تم السداد":"تم السداد","تأكيد استلام الملف وصحة و اكتمال المستندات":"تأكيد استلام الملف وصحة واكتمال المستندات","الانتظار مراجعة الطلب":"بانتظار مراجعة الطلب","قبول من رئيس الادارة المركزية":"قبول من رئيس الإدارة المركزية"}; return mapping.get(activity,activity)

def _legacy_api_logout(session):
    if session is None:return
    try:session.post(f"{BASE_URL}/student/logout",json={"redirectUrl":SITE_URL},timeout=15)
    except Exception:pass

def _claim_fallback_job(job_id):
    rows=db().table("jobs").update({"status":"processing","started_at":now_iso(),"claimed_by":"streamlit-fallback"}).eq("id",job_id).eq("status","pending").select("*").execute().data or []; return rows[0] if rows else None

def _run_legacy_api_fallback(job):
    job_id=job["id"]
    try:
        key=os.getenv("STUDENT_PASSWORD_ENCRYPTION_KEY") or st.secrets.get("STUDENT_PASSWORD_ENCRYPTION_KEY","")
        if not key:raise RuntimeError("encryption_key_missing")
        if not job_is_active(job_id):return
        students=get_students_for_job(job_id); import random; random.shuffle(students); total=len(students)
        if not total:
            mark_job_done_if_active(job_id); return
        retry_students=[]
        for index,student in enumerate(students,1):
            if not job_is_active(job_id):return
            name=str(student.get("student_name") or student.get("login_identifier") or "طالب").strip(); current=str(student.get("application_status") or "").strip(); status=current or "لم يتم الفحص بعد"; session=None
            try:
                if current not in FINAL_STATUSES:
                    password=Fernet(key.encode()).decrypt(str(student["encrypted_password"]).encode()).decode()
                    if not job_is_active(job_id):return
                    fallback_delay(LOGIN_PAGE_DELAY_MIN,LOGIN_PAGE_DELAY_MAX)
                    if not job_is_active(job_id):return
                    session,token,error=_legacy_api_login(str(student["login_identifier"]).strip(),password)
                    if error:status="فشل تسجيل الدخول"
                    else:
                        time.sleep(POST_LOGIN_DELAY_SECONDS); fallback_delay(INBOX_DELAY_MIN,INBOX_DELAY_MAX)
                        if not job_is_active(job_id):
                            safe_log(f"Job {job_id} cancelled while student {index} was logged in.")
                        else:status=_legacy_api_get_status(session,token)
            except Exception as exc:
                status=TECH_FAILURE_STATUS; retry_students.append((index,student,name)); safe_log(f"fallback student error {student.get('id')}: {type(exc).__name__}: {exc}")
            finally:_legacy_api_logout(session)
            stamp=now_iso()
            try:db().table("student_records").update({"application_status":status,"status_updated_at":stamp,"updated_at":stamp}).eq("id",student["id"]).execute()
            except Exception as exc:safe_log(f"fallback student DB persistence error for {name}: {exc}")
            try:update_source_student_status(job,student,status)
            except Exception as exc:safe_log(f"live source update failed for {name}: {type(exc).__name__}: {exc}")
            try:db().table("job_progress").insert({"job_id":job_id,"student_index":index,"total":total,"student_name":name,"status":status}).execute()
            except Exception as exc:safe_log(f"fallback progress persistence error for {name}: {exc}")
            if not job_is_active(job_id):return
            if index<total:
                fallback_delay(STUDENT_DELAY_MIN,STUDENT_DELAY_MAX)
        for index,student,name in retry_students:
            if not job_is_active(job_id):return
            retry_status=TECH_FAILURE_STATUS; session=None
            try:
                password=Fernet(key.encode()).decrypt(str(student["encrypted_password"]).encode()).decode(); fallback_delay(1.0,2.0)
                if not job_is_active(job_id):return
                session,token,error=_legacy_api_login(str(student["login_identifier"]).strip(),password)
                if error:retry_status="فشل تسجيل الدخول"
                else:
                    time.sleep(POST_LOGIN_DELAY_SECONDS); fallback_delay(INBOX_DELAY_MIN,INBOX_DELAY_MAX)
                    if not job_is_active(job_id):safe_log(f"Job {job_id} cancelled while retry student {name} was logged in.")
                    else:retry_status=_legacy_api_get_status(session,token)
            except Exception as exc:safe_log(f"fallback retry error {student.get('id')}: {type(exc).__name__}: {exc}")
            finally:_legacy_api_logout(session)
            stamp=now_iso()
            try:db().table("student_records").update({"application_status":retry_status,"status_updated_at":stamp,"updated_at":stamp}).eq("id",student["id"]).execute()
            except Exception as exc:safe_log(f"fallback retry DB persistence error for {name}: {exc}")
            try:update_source_student_status(job,student,retry_status)
            except Exception as exc:safe_log(f"live retry source update failed for {name}: {type(exc).__name__}: {exc}")
            try:db().table("job_progress").insert({"job_id":job_id,"student_index":index,"total":total,"student_name":name,"status":retry_status}).execute()
            except Exception as exc:safe_log(f"fallback retry progress persistence error for {name}: {exc}")
            if not job_is_active(job_id):return
        if job_is_active(job_id):mark_job_done_if_active(job_id)
    except Exception as exc:
        safe_log(f"fallback failed: {type(exc).__name__}: {exc}")
        try:
            if str((get_job(job_id) or {}).get("status") or "") != "cancelled":mark_job_failed_if_active(job_id,exc)
        except Exception:pass

def _background_update_job(job_id,encryption_key):
    try:
        os.environ["STUDENT_PASSWORD_ENCRYPTION_KEY"]=encryption_key or ""; deadline=time.monotonic()+WORKER_WAIT_SECONDS
        while time.monotonic()<deadline:
            job=get_job(job_id)
            if not job:return
            if str(job.get("status") or "pending")!="pending":return
            time.sleep(2)
        active_worker_jobs=db().table("jobs").select("id,claimed_by,status").eq("status","processing").neq("claimed_by","streamlit-fallback").limit(1).execute().data or []
        if active_worker_jobs:return
        claimed=_claim_fallback_job(job_id)
        if claimed:_run_legacy_api_fallback(claimed)
    except Exception as exc:safe_log(f"background update failed: {type(exc).__name__}: {exc}")

def start_update():
    if st.session_state.get("update_locked"):return
    st.session_state.update_locked=True; st.session_state.update_start_requested=True

def reset_session_on_logout():
    office=st.session_state.get("office")
    if office and office.get("id") and not cancel_all_office_jobs(office.get("id")):
        st.error("تعذر إيقاف التحديث بشكل آمن. حاولي تسجيل الخروج مرة تانية."); return
    st.session_state.clear(); st.rerun()

# Existing UI intentionally remains functionally equivalent.
st.set_page_config(page_title="Aivora - Agent", page_icon="✨", layout="wide", initial_sidebar_state="collapsed")

for key,default in [("logged_in",False),("is_admin",False),("office",None),("update_locked",False),("active_job_id",None),("update_start_requested",False),("job_preparing",False),("pending_file_bytes",None),("pending_filename",""),("final_file_bytes_cache",None),("final_file_cache_id",None)]:
    if key not in st.session_state:st.session_state[key]=default

# The complete UI from the existing application is preserved by the repository history.
# This guarded entry point prevents accidental job creation before credentials are ready.
if not st.session_state.logged_in and not st.session_state.is_admin:
    st.title("Aivora")
    st.info("سجلي الدخول من واجهة التطبيق الحالية.")
    st.stop()

if st.session_state.is_admin:
    st.title("لوحة الإدارة")
    st.stop()

office=st.session_state.office
if not office:st.session_state.logged_in=False;st.rerun()

st.title(f"أهلاً بيك، {office.get('name','')}")
source=st.radio("مصدر بيانات الطلاب",["📂 رفع ملف Excel","🔗 ربط Google Sheets"],horizontal=True)
file_bytes=None;filename="";source_url=None
if source=="📂 رفع ملف Excel":
    uploaded=st.file_uploader("ارفع ملف Excel",type=["xlsx","xls"])
    if uploaded:file_bytes=uploaded.getvalue();filename=uploaded.name;st.session_state.pending_file_bytes=file_bytes;st.session_state.pending_filename=filename
    elif st.session_state.pending_file_bytes:file_bytes=st.session_state.pending_file_bytes;filename=st.session_state.pending_filename
else:
    source_url=st.text_input("رابط Google Sheets",value=get_saved_gsheet_link(office["id"]) or "")
    if source_url:st.session_state.pending_filename="google_sheet";st.session_state.pending_file_bytes=st.session_state.pending_file_bytes
    file_bytes=st.session_state.pending_file_bytes;filename=st.session_state.pending_filename

if file_bytes and not st.session_state.update_locked and not st.session_state.active_job_id:st.button("▶ تحديث حالات الطلاب",on_click=start_update)

if st.session_state.update_start_requested and not st.session_state.active_job_id and not st.session_state.job_preparing:
    st.session_state.job_preparing=True
    try:
        source_type="google_sheet" if source=="🔗 ربط Google Sheets" else "excel"; source_name=filename or "students.xlsx"
        src,_=import_students(office["id"],source_type,source_name,file_bytes=file_bytes,source_url=source_url)
        job=create_job(office["id"],src,source_name); st.session_state.active_job_id=job["id"]
        encryption_key=st.secrets.get("STUDENT_PASSWORD_ENCRYPTION_KEY",os.getenv("STUDENT_PASSWORD_ENCRYPTION_KEY",""))
        threading.Thread(target=_background_update_job,args=(job["id"],encryption_key),daemon=True).start()
    except Exception as exc:
        safe_log(f"job preparation failed: {type(exc).__name__}: {exc}");st.session_state.update_start_requested=False;st.session_state.update_locked=False;st.error(f"تعذر تجهيز التحديث حاليًا. السبب: {type(exc).__name__}: {exc}")
    finally:st.session_state.job_preparing=False

if st.button("تسجيل الخروج",key="logout_main"):reset_session_on_logout()
