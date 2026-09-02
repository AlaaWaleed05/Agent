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

BASE_URL = "https://apiadm.study-in-egypt.gov.eg/api"
SITE_URL = "https://admission.study-in-egypt.gov.eg"
WORKER_WAIT_SECONDS = 30
# Keep fallback pacing aligned with the Worker timing policy.
LOGIN_PAGE_DELAY_MIN, LOGIN_PAGE_DELAY_MAX = 0.8, 1.5
POST_LOGIN_DELAY_SECONDS = 1.0
INBOX_DELAY_MIN, INBOX_DELAY_MAX = 0.7, 1.5
STUDENT_DELAY_MIN, STUDENT_DELAY_MAX = 4, 8
TECH_FAILURE_STATUS = "تعذر فحص الطالب حاليًا"
FINAL_STATUSES = {
    "مقبول نهائي", "قبول نهائي", "تم الرفض", "مرفوض نهائيًا",
    "مرفوض نهائيا", "مرفوض", "خالص",
}

ADMIN_USERNAME = st.secrets.get("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD", "admin123")


@st.cache_resource(show_spinner=False)
def get_supabase() -> Client:
    url = st.secrets.get("SUPABASE_URL", os.getenv("SUPABASE_URL"))
    key = st.secrets.get("SUPABASE_SERVICE_ROLE_KEY", os.getenv("SUPABASE_SERVICE_ROLE_KEY"))
    if not url or not key:
        raise RuntimeError("Supabase configuration is missing")
    return create_client(url, key)


def db():
    return get_supabase()


def now_iso():
    return datetime.now(timezone.utc).isoformat()


def safe_log(message):
    print(f"[Aivora] {message}")


def fallback_delay(a, b):
    """Use the same timing ranges as worker.py without touching Streamlit state."""
    import random
    time.sleep(random.uniform(a, b))


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
    if not name or not email or not password:
        return False, "كمّلي كل البيانات الأول."
    if len(password) < 6:
        return False, "الباسورد لازم يكون 6 أحرف على الأقل."
    try:
        if get_office_by_name(name) or get_office_by_email(email):
            return False, "اسم المكتب أو الإيميل مسجل قبل كده."
        password_hash = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
        db().table("offices").insert({"name": name, "email": email, "password_hash": password_hash, "status": "pending"}).execute()
        return True, "تم التسجيل بنجاح. الحساب في انتظار موافقة الإدارة."
    except Exception:
        safe_log("office registration failed")
        return False, "تعذر إنشاء الحساب حاليًا. حاولي مرة تانية."


def check_login(name, password):
    try:
        office = get_office_by_name(name)
        if not office:
            return False, "اسم المكتب أو الباسورد غلط!"
        if office.get("status") == "pending":
            return False, "حسابك في انتظار موافقة الإدارة!"
        if office.get("status") != "approved":
            return False, "حسابك موقوف!"
        full = get_office_with_hash(office["id"])
        password_hash = str((full or {}).get("password_hash") or "")
        if not password_hash or not bcrypt.checkpw(password.encode(), password_hash.encode()):
            return False, "اسم المكتب أو الباسورد غلط!"
        return True, office
    except Exception:
        safe_log("login check failed")
        return False, "تعذر تسجيل الدخول حاليًا. حاولي مرة تانية."


def get_pending_accounts():
    try:
        return db().table("offices").select("id,name,email,status,created_at").eq("status", "pending").order("created_at").execute().data or []
    except Exception:
        return []


def set_office_status(office_id, status):
    db().table("offices").update({"status": status}).eq("id", office_id).execute()


def log_activity(office_id, action, file_name="", details=None, student_id=None, data_source_id=None):
    try:
        db().table("activity_logs").insert({
            "office_id": office_id,
            "student_record_id": student_id,
            "data_source_id": data_source_id,
            "action": action,
            "file_name": file_name,
            "details": details or {},
        }).execute()
    except Exception as exc:
        safe_log(f"activity log error: {exc}")


def get_gsheet_client():
    if gspread is None or Credentials is None:
        raise RuntimeError("Google Sheets libraries unavailable")
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"],
    )
    return gspread.authorize(creds)


def extract_sheet_id(link):
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", str(link))
    return match.group(1) if match else None


def extract_gid(link):
    match = re.search(r"[?#&]gid=(\d+)", str(link))
    return int(match.group(1)) if match else None


def read_gsheet_rows(link):
    sheet_id = extract_sheet_id(link)
    if not sheet_id:
        raise ValueError("رابط Google Sheets غير صحيح.")
    spreadsheet = get_gsheet_client().open_by_key(sheet_id)
    gid = extract_gid(link)
    worksheet = next((w for w in spreadsheet.worksheets() if w.id == gid), spreadsheet.sheet1) if gid is not None else spreadsheet.sheet1
    return worksheet.get_all_values()


def get_saved_gsheet_link(office_id):
    try:
        rows = db().table("data_sources").select("source_url,created_at").eq("office_id", office_id).eq("source_type", "google_sheet").not_.is_("source_url", "null").order("created_at", desc=True).limit(1).execute().data or []
        return rows[0].get("source_url") if rows else None
    except Exception:
        return None


def save_gsheet_link(office_id, link):
    try:
        if not extract_sheet_id(link):
            return False, "الرابط غير صحيح!"
        db().table("data_sources").insert({
            "office_id": office_id,
            "source_type": "google_sheet",
            "source_name": "Google Sheet",
            "source_url": link,
            "column_mapping": {},
        }).execute()
        return True, "تم حفظ الرابط بنجاح"
    except Exception:
        safe_log("Google Sheet link save failed")
        return False, "تعذر حفظ الرابط حاليًا."


def find_excel_columns(ws):
    cols = {"name": None, "email": None, "password": None}
    header_row = None
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), start=1):
        values = [str(c).strip() if c is not None else "" for c in row]
        if any("يميل" in v or "mail" in v.lower() or "بريد" in v for v in values):
            header_row = row_idx
            for i, cell in enumerate(values):
                low = cell.lower()
                if any(k in cell for k in ["اسم", "الإسم", "الاسم"]) or "name" in low:
                    cols["name"] = i
                elif any(k in cell for k in ["يميل", "بريد"]) or "mail" in low:
                    cols["email"] = i
                elif any(k in cell for k in ["باسورد", "كلمة المرور", "password", "pass"]):
                    cols["password"] = i
            break
    if header_row is None:
        raise ValueError("مش لاقي هيدر الإكسيل.")
    if cols["email"] is None:
        raise ValueError("مش لاقي عمود الإيميل.")
    if cols["password"] is None:
        raise ValueError("مش لاقي عمود الباسورد.")
    if cols["name"] is None:
        cols["name"] = cols["email"]
    return cols, header_row


def parse_excel_bytes(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
    ws = wb.active
    cols, header_row = find_excel_columns(ws)
    records = []
    seen = set()
    for excel_row, row in enumerate(ws.iter_rows(min_row=header_row + 1, values_only=True), start=header_row + 1):
        values = list(row)
        email = str(values[cols["email"]] or "").strip() if cols["email"] < len(values) else ""
        password = str(values[cols["password"]] or "").strip() if cols["password"] < len(values) else ""
        name = str(values[cols["name"]] or "").strip() if cols["name"] < len(values) else email
        key = email.lower()
        if not email or not password or key in seen:
            continue
        seen.add(key)
        records.append({
            "source_row_number": excel_row,
            "student_name": name or email,
            "login_identifier": email,
            "password": password,
            "original_data": {f"column_{i+1}": (str(v) if v is not None else "") for i, v in enumerate(values)},
        })
    return records


def encrypt_password(password, key):
    if not key:
        raise RuntimeError("Encryption key missing")
    return Fernet(key.encode()).encrypt(password.encode()).decode()


def import_students(office_id, source_type, source_name, file_bytes=None, source_url=None):
    if source_type in {"xlsx", "xls", "excel"}:
        records = parse_excel_bytes(file_bytes)
    else:
        rows = read_gsheet_rows(source_url)
        if not rows:
            raise ValueError("الشيت فاضي.")
        wb = openpyxl.Workbook()
        ws = wb.active
        for row in rows:
            ws.append(row)
        out = io.BytesIO()
        wb.save(out)
        records = parse_excel_bytes(out.getvalue())
    if not records:
        raise ValueError("مش لاقي طلاب عندهم إيميل وباسورد صالحين.")

    encryption_key = st.secrets.get("STUDENT_PASSWORD_ENCRYPTION_KEY", os.getenv("STUDENT_PASSWORD_ENCRYPTION_KEY"))
    source_type = "google_sheet" if source_type in {"gsheet", "google_sheet"} else "excel"
    source = db().table("data_sources").insert({
        "office_id": office_id,
        "source_type": source_type,
        "source_name": source_name,
        "source_url": source_url,
        "column_mapping": {},
    }).execute().data[0]

    payload = [{
        "office_id": office_id,
        "data_source_id": source["id"],
        "source_row_number": r["source_row_number"],
        "student_name": r["student_name"],
        "login_identifier": r["login_identifier"],
        "encrypted_password": encrypt_password(r["password"], encryption_key),
        "application_status": "",
        "original_data": r["original_data"],
        "updated_at": now_iso(),
    } for r in records]
    db().table("student_records").insert(payload).execute()
    return source, len(payload)


def create_job(office_id, source, file_name):
    return db().table("jobs").insert({
        "office_id": office_id,
        "data_source_id": source["id"],
        "source_type": source["source_type"],
        "source_ref": source.get("source_url") or source["id"],
        "file_name": file_name,
        "status": "pending",
    }).execute().data[0]


def get_job(job_id):
    if not job_id:
        return None
    rows = db().table("jobs").select("*").eq("id", job_id).limit(1).execute().data or []
    return rows[0] if rows else None


def get_job_progress_rows(job_id):
    rows = db().table("job_progress").select("student_index,total,student_name,status,created_at").eq("job_id", job_id).order("created_at").execute().data or []
    latest = {}
    for row in rows:
        key = str(row.get("student_name") or "").strip().lower()
        if key:
            latest[key] = row
    return sorted(latest.values(), key=lambda r: int(r.get("student_index") or 0))


def get_students_for_job(job_id):
    job = get_job(job_id)
    if not job:
        return []
    rows = db().table("student_records").select("id,student_name,login_identifier,application_status,status_updated_at,source_row_number,created_at,updated_at").eq("office_id", job["office_id"]).eq("data_source_id", job["data_source_id"]).order("source_row_number").execute().data or []
    latest = {}
    for row in rows:
        key = str(row.get("login_identifier") or row.get("student_name") or "").strip().lower()
        if not key:
            continue
        latest[key] = row
    return list(latest.values())


def get_students(office_id, search=""):
    rows = db().table("student_records").select("id,student_name,login_identifier,application_status,status_updated_at,source_row_number,created_at,updated_at").eq("office_id", office_id).execute().data or []
    def stamp(v):
        try:
            return datetime.fromisoformat(str(v or "").replace("Z", "+00:00"))
        except Exception:
            return datetime.min.replace(tzinfo=timezone.utc)
    latest = {}
    for row in rows:
        key = str(row.get("login_identifier") or row.get("student_name") or "").strip().lower()
        if not key:
            continue
        score = (stamp(row.get("status_updated_at")), stamp(row.get("updated_at")), stamp(row.get("created_at")))
        if key not in latest or score > latest[key][0]:
            latest[key] = (score, row)
    rows = sorted([x[1] for x in latest.values()], key=lambda r: str(r.get("student_name") or "").lower())
    q = search.strip().lower()
    return [r for r in rows if not q or q in str(r.get("student_name") or "").lower()]


def _legacy_api_login(email, password):
    session = requests.Session()
    session.headers.update({
        "accept": "application/json, text/plain, */*",
        "accept-language": "ar",
        "device": "CITIZEN",
        "origin": SITE_URL,
        "referer": SITE_URL + "/",
        "user-agent": "Mozilla/5.0",
        "content-type": "application/json",
    })
    response = session.post(f"{BASE_URL}/student/login", json={"email": email, "password": password}, timeout=30)
    if response.status_code not in (200, 201):
        return None, None, "login_failed"
    body = response.json() if response.content else {}
    return session, body.get("token", "") or response.headers.get("x-csrf-token", ""), None


def _legacy_api_get_status(session, token):
    filt = {"where": {}, "limit": 10, "offset": 0, "order": "statusUpdatedAt DESC", "fields": ["serviceSlug", "ID", "createdAt", "statusUpdatedAt", "activityId", "activityName"]}
    headers = {"x-csrf-token": token} if token else {}
    response = session.get(f"{BASE_URL}/dynamic_services/inbox", params={"filter": json.dumps(filt)}, headers=headers, timeout=30)
    if response.status_code not in (200, 304):
        raise RuntimeError(f"status_http_{response.status_code}")
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
    rows = db().table("jobs").update({"status": "processing", "started_at": now_iso(), "claimed_by": "streamlit-fallback"}).eq("id", job_id).eq("status", "pending").select("*").execute().data or []
    return rows[0] if rows else None


def _run_legacy_api_fallback(job_id):
    """Old API path with the same pacing policy used by the Worker."""
    try:
        client = db()
        key = os.getenv("STUDENT_PASSWORD_ENCRYPTION_KEY")
        if not key:
            raise RuntimeError("encryption_key_missing")
        students = client.table("student_records").select("*").eq("office_id", job_id["office_id"]).eq("data_source_id", job_id["data_source_id"]).order("source_row_number").execute().data or []
        unique = {}
        for student in students:
            key_id = str(student.get("login_identifier") or student.get("student_name") or "").strip().lower()
            if key_id:
                unique[key_id] = student
        students = list(unique.values())
        import random
        random.shuffle(students)
        total = len(students)
        if not total:
            client.table("jobs").update({"status": "done", "finished_at": now_iso()}).eq("id", job_id["id"]).execute()
            return

        for index, student in enumerate(students, 1):
            name = str(student.get("student_name") or student.get("login_identifier") or "طالب").strip()
            current = str(student.get("application_status") or "").strip()
            status = current or "لم يتم الفحص بعد"
            session = None
            try:
                if current in FINAL_STATUSES:
                    status = current
                else:
                    password = Fernet(key.encode()).decrypt(str(student["encrypted_password"]).encode()).decode()

                    # Same delay used by the Worker after opening the login page.
                    fallback_delay(LOGIN_PAGE_DELAY_MIN, LOGIN_PAGE_DELAY_MAX)
                    session, token, error = _legacy_api_login(str(student["login_identifier"]).strip(), password)
                    if error:
                        status = "فشل تسجيل الدخول"
                    else:
                        # Same login-transition wait used by the Worker before continuing.
                        time.sleep(POST_LOGIN_DELAY_SECONDS)
                        # Same inbox-navigation delay used by the Worker.
                        fallback_delay(INBOX_DELAY_MIN, INBOX_DELAY_MAX)
                        status = _legacy_api_get_status(session, token)
            except Exception as exc:
                status = TECH_FAILURE_STATUS
                safe_log(f"fallback student error {student.get('id')}: {type(exc).__name__}: {exc}")
            finally:
                _legacy_api_logout(session)
                if session is not None:
                    # Same short pause used by the Worker after logout.
                    fallback_delay(INBOX_DELAY_MIN, INBOX_DELAY_MAX)

            stamp = now_iso()
            try:
                client.table("student_records").update({"application_status": status, "status_updated_at": stamp, "updated_at": stamp}).eq("id", student["id"]).execute()
                client.table("job_progress").insert({"job_id": job_id["id"], "student_index": index, "total": total, "student_name": name, "status": status}).execute()
            except Exception as exc:
                safe_log(f"fallback persistence error for {name}: {exc}")

            # Exactly the same 4–8 second student-to-student delay as the Worker.
            if index < total:
                fallback_delay(STUDENT_DELAY_MIN, STUDENT_DELAY_MAX)

        client.table("jobs").update({"status": "done", "finished_at": now_iso(), "error": None}).eq("id", job_id["id"]).execute()
    except Exception as exc:
        safe_log(f"fallback failed: {type(exc).__name__}: {exc}")
        try:
            db().table("jobs").update({"status": "failed", "finished_at": now_iso(), "error": str(exc)[:1000]}).eq("id", job_id["id"]).execute()
        except Exception as db_exc:
            safe_log(f"fallback job finalization failed: {db_exc}")

def _background_update_job(job_id, encryption_key):
    try:
        # Keep all Streamlit state/UI calls out of the background thread.
        os.environ["STUDENT_PASSWORD_ENCRYPTION_KEY"] = encryption_key or ""
        deadline = time.monotonic() + WORKER_WAIT_SECONDS
        while time.monotonic() < deadline:
            job = get_job(job_id)
            if not job:
                return
            if str(job.get("status") or "pending") != "pending":
                return
            time.sleep(2)
        claimed = _claim_fallback_job(job_id)
        if claimed:
            _run_legacy_api_fallback(claimed)
    except Exception as exc:
        safe_log(f"background update failed: {type(exc).__name__}: {exc}")
        try:
            db().table("jobs").update({"status": "failed", "finished_at": now_iso(), "error": str(exc)[:1000]}).eq("id", job_id).execute()
        except Exception as db_exc:
            safe_log(f"background finalization failed: {db_exc}")


def start_update():
    if st.session_state.get("update_locked"):
        return
    st.session_state.update_locked = True
    st.session_state.update_start_requested = True


def reset_session_on_logout():
    st.session_state.clear()
    st.rerun()


st.set_page_config(page_title="Aivora - Agent", page_icon="✨", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
html, body, .stApp, [class*="css"] { font-family: 'Cairo', sans-serif !important; direction: rtl; }
.stApp { background: #f6f8fc; color: #111827; }
.block-container { max-width: 1120px; padding-top: 1.25rem; padding-bottom: 2rem; }
#MainMenu, footer, header, [data-testid="stToolbar"], [data-testid="stDecoration"], [data-testid="stStatusWidget"], [data-testid="stSidebarNav"] { display:none !important; }
.card { background:#fff; border:1px solid #e5e7eb; border-radius:18px; padding:20px; margin-bottom:16px; box-shadow:0 3px 14px rgba(17,24,39,.04); }
.topbar { background:#fff; border:1px solid #e5e7eb; border-radius:16px; padding:14px 18px; margin-bottom:18px; }
.brand-title { font-size:19px; font-weight:800; color:#111827; }
.brand-sub { font-size:12px; color:#6b7280; margin-top:2px; }
.section-title { font-size:18px; font-weight:800; margin:2px 0 10px; }
.section-sub { color:#6b7280; font-size:13px; margin-bottom:14px; }
.stButton > button { width:100%; min-height:44px; border-radius:11px !important; font-weight:700 !important; }
.connected-box { background:#f0fdf4; border:1px solid #bbf7d0; border-radius:11px; padding:11px 14px; color:#166534; font-size:13px; margin-bottom:12px; }
.wait-box { background:#eff6ff; border:1px solid #bfdbfe; border-radius:12px; padding:13px 15px; color:#1d4ed8; font-weight:700; }
.success-box { background:#ecfdf5; border:1px solid #bbf7d0; border-radius:12px; padding:14px 15px; color:#166534; font-weight:800; }
.lock-box { background:#fff7ed; border:1px solid #fed7aa; border-radius:12px; padding:13px 15px; color:#9a3412; }
.progress-table { width:100%; border-collapse:separate; border-spacing:0; overflow:hidden; border:1px solid #e5e7eb; border-radius:12px; }
.progress-table th, .progress-table td { padding:10px 12px; border-bottom:1px solid #eef0f4; font-size:13px; }
.progress-table th { background:#f8fafc; color:#6b7280; font-weight:800; }
.progress-table tr:last-child td { border-bottom:0; }
.progress-table .num { width:48px; text-align:center; direction:ltr; }
.progress-table .status { text-align:left; direction:rtl; }
.progress-table .name { text-align:right; }
[data-testid="InputInstructions"] { display:none !important; }
[data-testid="stTextInput"] button { display:none !important; }
</style>
""", unsafe_allow_html=True)

for key, default in [
    ("logged_in", False), ("is_admin", False), ("office", None),
    ("update_locked", False), ("active_job_id", None),
    ("update_start_requested", False), ("job_preparing", False),
    ("pending_file_bytes", None), ("pending_filename", ""),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# Login
if not st.session_state.logged_in and not st.session_state.is_admin:
    st.markdown("<div style='text-align:center;margin:36px 0 22px'><div style='font-size:46px'>✨</div><div style='font-size:31px;font-weight:800'>Aivora</div><div style='font-size:14px;color:#6b7280'>Your Smarter Support for Every Student's Application</div></div>", unsafe_allow_html=True)
    left, right = st.columns([1.15, 1], gap="large")
    with left:
        st.markdown("<div style='padding:34px 18px'><div style='font-size:14px;color:#2563eb;font-weight:800'>حل بسيط لإدارة المتابعة</div><div style='font-size:30px;font-weight:800;line-height:1.35'>تابع طلبات طلابك<br>من مكان واحد.</div><div style='font-size:15px;color:#6b7280;line-height:1.9;margin-top:12px'>حدّث حالات الطلبات، اربط Google Sheets، وابحث عن أي طالب بسرعة بدون متابعة يدوية.</div></div>", unsafe_allow_html=True)
    with right:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        tab1, tab2 = st.tabs(["تسجيل الدخول", "حساب جديد"])
        with tab1:
            username = st.text_input("اسم المكتب", key="login_user")
            password = st.text_input("كلمة المرور", type="password", key="login_pass")
            if st.button("تسجيل الدخول", key="login_btn"):
                if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
                    st.session_state.is_admin = True
                    st.rerun()
                else:
                    ok, result = check_login(username, password)
                    if ok:
                        st.session_state.logged_in = True
                        st.session_state.office = result
                        st.session_state.update_locked = False
                        st.rerun()
                    st.error(result)
        with tab2:
            new_office = st.text_input("اسم المكتب", key="reg_office")
            new_email = st.text_input("الإيميل", key="reg_email")
            new_pass = st.text_input("كلمة المرور", type="password", key="reg_pass")
            new_pass2 = st.text_input("تأكيد كلمة المرور", type="password", key="reg_pass2")
            if st.button("إنشاء الحساب", key="reg_btn"):
                if new_pass != new_pass2:
                    st.error("كلمة المرور مش متطابقة!")
                else:
                    ok, msg = register_office(new_office, new_email, new_pass)
                    (st.success if ok else st.error)(msg)
        st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

# Admin
if st.session_state.is_admin:
    st.markdown('<div class="card"><div class="section-title">الإدارة</div></div>', unsafe_allow_html=True)
    for acc in get_pending_accounts():
        c1, c2, c3, c4 = st.columns([3, 3, 1, 1])
        c1.write(acc.get("name", "")); c2.write(acc.get("email", ""))
        if c3.button("قبول", key=f"approve_{acc['id']}"):
            set_office_status(acc["id"], "approved"); st.rerun()
        if c4.button("رفض", key=f"reject_{acc['id']}"):
            set_office_status(acc["id"], "rejected"); st.rerun()
    if st.button("تسجيل الخروج من الإدارة", key="admin_logout"):
        reset_session_on_logout()
    st.stop()

office = st.session_state.office
if not office:
    st.session_state.logged_in = False
    st.rerun()
office_id = office["id"]

st.markdown("<div class='topbar'><div class='brand-title'>Aivora</div><div class='brand-sub'>Your Smarter Support for Every Student's Application</div></div>", unsafe_allow_html=True)

# Source selection
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown('<div class="section-title">مصدر بيانات الطلاب</div><div class="section-sub">اختاري المصدر الذي يحتوي على بيانات الطلاب.</div>', unsafe_allow_html=True)
source = st.radio("", ["📂 رفع ملف Excel", "🔗 ربط Google Sheets"], horizontal=True, label_visibility="collapsed", key="source_mode")
saved_link = get_saved_gsheet_link(office_id)
file_bytes = None
filename = ""
source_url = None

if source == "📂 رفع ملف Excel":
    uploaded = st.file_uploader("ارفع ملف Excel", type=["xlsx", "xls"], label_visibility="collapsed", key="excel_upload")
    if uploaded:
        file_bytes = uploaded.getvalue()
        filename = uploaded.name
        st.session_state.pending_file_bytes = file_bytes
        st.session_state.pending_filename = filename
        st.success(f"تم اختيار الملف: {filename}")
    elif st.session_state.pending_file_bytes:
        file_bytes = st.session_state.pending_file_bytes
        filename = st.session_state.pending_filename
else:
    if saved_link:
        st.markdown('<div class="connected-box">✓ Google Sheets متصل بالفعل لهذا المكتب</div>', unsafe_allow_html=True)
        change = st.checkbox("تغيير رابط الشيت", key="change_link")
        if change:
            new_link = st.text_input("رابط Google Sheets", key="new_link")
            if st.button("حفظ الرابط", key="save_link"):
                ok, msg = save_gsheet_link(office_id, new_link)
                (st.success if ok else st.error)(msg)
                if ok: st.rerun()
        source_url = saved_link
    else:
        new_link = st.text_input("رابط Google Sheets", key="first_link")
        if st.button("حفظ وربط الشيت", key="save_first_link"):
            ok, msg = save_gsheet_link(office_id, new_link)
            (st.success if ok else st.error)(msg)
            if ok: st.rerun()
    if source_url and st.button("تحميل بيانات الشيت", key="load_sheet"):
        try:
            rows = read_gsheet_rows(source_url)
            wb = openpyxl.Workbook(); ws = wb.active
            for row in rows: ws.append(row)
            out = io.BytesIO(); wb.save(out)
            st.session_state.pending_file_bytes = out.getvalue()
            st.session_state.pending_filename = "google_sheet"
            st.success("تم جلب البيانات. جاهزة للتحديث.")
        except Exception:
            st.error("تعذر جلب بيانات الشيت حاليًا.")
    file_bytes = st.session_state.pending_file_bytes
    filename = st.session_state.pending_filename

st.markdown('</div>', unsafe_allow_html=True)

# Update request and preparation happen on the main Streamlit thread.
if file_bytes and not st.session_state.update_locked and not st.session_state.active_job_id:
    st.button("▶ تحديث حالات الطلاب", key="start_update", on_click=start_update)

if st.session_state.update_start_requested and not st.session_state.active_job_id and not st.session_state.job_preparing:
    running = db().table("jobs").select("id").eq("office_id", office_id).in_("status", ["pending", "processing"]).limit(1).execute().data or []
    if running:
        st.session_state.update_start_requested = False
        st.session_state.update_locked = True
        st.warning("في تحديث شغال بالفعل لهذا المكتب. استني لحد ما يخلص.")
    else:
        st.session_state.job_preparing = True
        try:
            if source == "🔗 ربط Google Sheets":
                source_type = "google_sheet"
                source_name = "Google Sheet"
            else:
                source_type = "excel"
                source_name = filename or "students.xlsx"
            src, _ = import_students(office_id, source_type, source_name, file_bytes=file_bytes, source_url=source_url)
            job = create_job(office_id, src, source_name)
            st.session_state.active_job_id = job["id"]
            encryption_key = st.secrets.get("STUDENT_PASSWORD_ENCRYPTION_KEY", os.getenv("STUDENT_PASSWORD_ENCRYPTION_KEY", ""))
            threading.Thread(target=_background_update_job, args=(job["id"], encryption_key), daemon=True).start()
        except Exception as exc:
            safe_log(f"job preparation failed: {type(exc).__name__}: {exc}")
            st.session_state.update_start_requested = False
            st.session_state.update_locked = False
            st.error("تعذر تجهيز التحديث حاليًا. حاولي مرة تانية.")
        finally:
            st.session_state.job_preparing = False

@st.fragment(run_every=2)
def render_processing():
    job_id = st.session_state.get("active_job_id")
    if not job_id and not st.session_state.get("update_locked"):
        return

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">تحديث حالات الطلاب</div>', unsafe_allow_html=True)
    job = get_job(job_id) if job_id else None
    if not job:
        if st.session_state.get("update_locked"):
            st.markdown('<div class="lock-box">🔒 تم تشغيل تحديث بالفعل في هذه الجلسة. لو عايزة تبدئي تحديث جديد، سجّلي خروج وادخلي تاني.</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        return

    status = str(job.get("status") or "pending")
    if status == "pending":
        st.markdown('<div class="wait-box">⏳ سيبدأ التحديث خلال ثواني…</div>', unsafe_allow_html=True)
    elif status == "processing":
        st.markdown('<div class="wait-box">▶️ بدأ التحديث. جاري فحص الطلاب وتحديث الحالات…</div>', unsafe_allow_html=True)

    students = get_students_for_job(job_id)
    progress_rows = get_job_progress_rows(job_id)
    progress_map = {str(r.get("student_name") or "").strip().lower(): r for r in progress_rows}

    if students:
        table_rows = []
        checked = 0
        for index, student in enumerate(students, 1):
            name = str(student.get("student_name") or student.get("login_identifier") or "طالب")
            row = progress_map.get(name.strip().lower())
            current = row.get("status") if row else str(student.get("application_status") or "لم يتم الفحص بعد")
            if row: checked += 1
            table_rows.append((index, name, current))
        total = len(table_rows)
        st.progress(min(checked / max(total, 1), 1.0))
        st.caption(f"تم فحص {min(checked, total)} من {total} طالب")
        if progress_rows:
            last = progress_rows[-1]
            st.info(f"🔄 آخر طالب تم فحصه: **{last.get('student_name') or 'طالب'}** — الحالة: **{last.get('status') or ''}**")
        html = ['<table class="progress-table"><thead><tr><th class="num">#</th><th class="name">اسم الطالب</th><th class="status">الحالة</th></tr></thead><tbody>']
        for number, name, current in table_rows:
            html.append(f'<tr><td class="num">{number}</td><td class="name">{name}</td><td class="status">{current}</td></tr>')
        html.append('</tbody></table>')
        st.markdown(''.join(html), unsafe_allow_html=True)

    if status == "done":
        st.markdown('<div class="success-box" style="margin-top:14px">اكتمل التحديث 🎉</div>', unsafe_allow_html=True)
    elif status == "failed":
        # Technical details remain in Supabase logs/job.error; office sees only this safe message.
        st.info("تعذر إكمال التحديث حاليًا. حاولي مرة تانية بعد قليل.")
    elif not job_id and st.session_state.get("update_locked"):
        st.markdown('<div class="lock-box">🔒 تم تشغيل تحديث بالفعل في هذه الجلسة. لو عايزة تبدئي تحديث جديد، سجّلي خروج وادخلي تاني.</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

render_processing()

# Search
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown('<div class="section-title">البحث عن طالب</div>', unsafe_allow_html=True)
search_query = st.text_input("اسم الطالب", label_visibility="collapsed", key="student_search")
if search_query:
    found = get_students(office_id, search_query)
    if found:
        for student in found:
            status = student.get("application_status") or "لم يتم الفحص بعد"
            st.markdown(f'<div style="background:#f8fafc;border:1px solid #e5e7eb;border-radius:11px;padding:12px;margin:7px 0"><b>👤 {student.get("student_name", "")}</b><br><span style="color:#1d4ed8;font-size:13px">{status}</span></div>', unsafe_allow_html=True)
    else:
        st.info("مفيش طالب بالاسم ده.")
st.markdown('</div>', unsafe_allow_html=True)

# Logout is intentionally at the bottom. Logging out clears the session lock.
if st.button("تسجيل الخروج", key="logout_main"):
    reset_session_on_logout()
