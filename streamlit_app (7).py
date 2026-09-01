"""
Aivora - Agent
Your Smarter Support for Every Student's Application
Streamlit App

Supabase-backed web app. Google Sheets remains an optional INPUT source;
accounts, students, jobs, progress and activity logs are stored in Supabase.
"""

import io
import os
import re
import time
from datetime import datetime, timezone

import bcrypt
import openpyxl
import pandas as pd
import plotly.express as px
import streamlit as st
from cryptography.fernet import Fernet
from supabase import create_client, Client

try:
    import gspread
    from google.oauth2.service_account import Credentials
except Exception:
    gspread = None
    Credentials = None


# ==================== إعدادات ====================
ADMIN_USERNAME = st.secrets.get("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD", "admin123")

FINAL_STATUSES = {
    "مقبول نهائي", "قبول نهائي", "تم الرفض", "مرفوض نهائيًا",
    "مرفوض نهائيا", "مرفوض", "خالص",
}


# ==================== Supabase ====================
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
    rows = (
        db().table("offices")
        .select("id,name,email,status,created_at")
        .eq("name", str(name).strip())
        .limit(1).execute().data or []
    )
    return rows[0] if rows else None


def get_office_with_hash(office_id):
    rows = (
        db().table("offices")
        .select("id,name,email,status,password_hash,created_at")
        .eq("id", office_id).limit(1).execute().data or []
    )
    return rows[0] if rows else None


def get_office_by_email(email):
    rows = (
        db().table("offices")
        .select("id,name,email,status")
        .eq("email", str(email).strip().lower())
        .limit(1).execute().data or []
    )
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
        db().table("offices").insert({
            "name": name,
            "email": email,
            "password_hash": password_hash,
            "status": "pending",
        }).execute()
        return True, "تم التسجيل بنجاح. الحساب في انتظار موافقة الإدارة."
    except Exception as exc:
        return False, f"حصل خطأ أثناء التسجيل: {exc}"


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
    except Exception as exc:
        return False, f"حصل خطأ في تسجيل الدخول: {exc}"


def get_pending_accounts():
    try:
        return (
            db().table("offices")
            .select("id,name,email,status,created_at")
            .eq("status", "pending")
            .order("created_at")
            .execute().data or []
        )
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
        print(f"Activity log error: {exc}")


# ==================== Google Sheet input ====================
def get_gsheet_client():
    if gspread is None or Credentials is None:
        raise RuntimeError("مكتبات Google Sheets مش مثبتة.")
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"],
    )
    return gspread.authorize(creds)


def extract_sheet_id(link):
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", str(link))
    return m.group(1) if m else None


def extract_gid(link):
    m = re.search(r"[?#&]gid=(\d+)", str(link))
    return int(m.group(1)) if m else None


def read_gsheet_rows(link):
    sheet_id = extract_sheet_id(link)
    if not sheet_id:
        raise ValueError("رابط Google Sheets غير صحيح.")
    spreadsheet = get_gsheet_client().open_by_key(sheet_id)
    gid = extract_gid(link)
    ws = next((w for w in spreadsheet.worksheets() if w.id == gid), spreadsheet.sheet1) if gid is not None else spreadsheet.sheet1
    return ws.get_all_values()


# ==================== Excel / student import ====================
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
    for excel_row, row in enumerate(ws.iter_rows(min_row=header_row + 1, values_only=True), start=header_row + 1):
        values = list(row)
        email = str(values[cols["email"]] or "").strip() if cols["email"] < len(values) else ""
        password = str(values[cols["password"]] or "").strip() if cols["password"] < len(values) else ""
        name = str(values[cols["name"]] or "").strip() if cols["name"] < len(values) else email
        if not email or not password:
            continue
        records.append({
            "source_row_number": excel_row,
            "student_name": name or email,
            "login_identifier": email,
            "password": password,
            "original_data": {f"column_{i+1}": (str(v) if v is not None else "") for i, v in enumerate(values)},
        })
    return records


def encrypt_password(password):
    key = st.secrets.get("STUDENT_PASSWORD_ENCRYPTION_KEY", os.getenv("STUDENT_PASSWORD_ENCRYPTION_KEY"))
    if not key:
        raise RuntimeError("STUDENT_PASSWORD_ENCRYPTION_KEY مش موجود في Secrets.")
    return Fernet(key.encode()).encrypt(password.encode()).decode()


def upsert_students(office_id, source_id, records):
    payload = []
    for r in records:
        payload.append({
            "office_id": office_id,
            "data_source_id": source_id,
            "source_row_number": r["source_row_number"],
            "student_name": r["student_name"],
            "login_identifier": r["login_identifier"],
            "encrypted_password": encrypt_password(r["password"]),
            "application_status": "",
            "original_data": r["original_data"],
            "updated_at": now_iso(),
        })
    db().table("student_records").insert(payload).execute()
    return len(payload)


def create_data_source(office_id, source_type, source_name, source_url=None, mapping=None):
    source_type_map = {
        "xlsx": "excel",
        "xls": "excel",
        "gsheet": "google_sheet",
        "google_sheet": "google_sheet",
        "csv": "csv",
    }

    normalized_source_type = source_type_map.get(source_type, source_type)

    row = db().table("data_sources").insert({
        "office_id": office_id,
        "source_type": normalized_source_type,
        "source_name": source_name,
        "source_url": source_url,
        "column_mapping": mapping or {},
    }).execute().data

    return row[0]

def import_students(office_id, source_type, source_name, file_bytes=None, source_url=None):
    if source_type == "xlsx":
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
    source = create_data_source(office_id, source_type, source_name, source_url=source_url)
    count = upsert_students(office_id, source["id"], records)
    log_activity(office_id, "استيراد بيانات الطلاب إلى Supabase", source_name, {"students": count, "source_type": source_type}, data_source_id=source["id"])
    return source, count


# ==================== Jobs ====================
def create_job(office_id, source, file_name):
    rows = db().table("jobs").insert({
        "office_id": office_id,
        "data_source_id": source["id"],
        "source_type": source["source_type"],
        "source_ref": source.get("source_url") or source["id"],
        "file_name": file_name,
        "status": "pending",
    }).execute().data
    return rows[0]


def get_job(job_id):
    rows = db().table("jobs").select("*").eq("id", job_id).limit(1).execute().data or []
    return rows[0] if rows else None


def get_job_progress(job_id):
    return (
        db().table("job_progress")
        .select("student_index,total,student_name,status,created_at")
        .eq("job_id", job_id)
        .order("student_index", desc=True).limit(1)
        .execute().data or []
    )


def get_students(office_id, search=""):
    rows = (
        db().table("student_records")
        .select("id,student_name,login_identifier,application_status,status_updated_at,source_row_number")
        .eq("office_id", office_id)
        .order("student_name")
        .execute().data or []
    )
    if search.strip():
        q = search.strip().lower()
        rows = [r for r in rows if q in str(r.get("student_name", "")).lower()]
    return rows


def status_class(status):
    s = str(status or "")
    if any(x in s for x in ["مقبول", "خالص"]): return "status-ok"
    if any(x in s for x in ["رفض", "فشل", "خطأ"]): return "status-error"
    if any(x in s for x in ["مفيش", "انتظار", "مراجعة"]): return "status-warn"
    return "status-info"


# ==================== UI ====================
st.set_page_config(page_title="Aivora - Agent", page_icon="✨", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800&display=swap');
html,body,[class*="css"],.stApp{font-family:'Cairo',sans-serif!important;direction:rtl}.stApp{background:#f5f7fb;color:#111827}.block-container{max-width:1180px;padding-top:1.4rem;padding-bottom:3rem}#MainMenu,footer,header,[data-testid="stToolbar"],[data-testid="stDecoration"],[data-testid="stStatusWidget"]{display:none!important}h1,h2,h3,p,label,span,div{font-family:'Cairo',sans-serif!important}h1{font-size:32px!important;font-weight:800!important;color:#111827!important}.topbar,.card,.stat-card{background:#fff;border:1px solid #e5e7eb;border-radius:16px;box-shadow:0 3px 14px rgba(17,24,39,.045)}.topbar{padding:13px 18px;margin-bottom:22px;display:flex;justify-content:space-between;align-items:center}.card{padding:22px;margin-bottom:18px}.brand{display:flex;align-items:center;gap:11px}.brand-icon{width:42px;height:42px;border-radius:12px;background:#eff6ff;display:flex;align-items:center;justify-content:center;font-size:22px}.brand-title{font-size:18px;font-weight:800}.brand-sub{font-size:12px;color:#6b7280}.hero{background:linear-gradient(135deg,#fff,#f8fbff);border:1px solid #dbeafe;border-radius:20px;padding:25px 28px;margin-bottom:20px}.hero-kicker{color:#6b7280;font-size:14px;font-weight:600}.hero-title{font-size:28px;font-weight:800}.hero-title strong{color:#2563eb}.hero-desc{color:#6b7280;font-size:14px}.stat-card{padding:17px 18px}.stat-label{color:#6b7280;font-size:12px}.stat-value{font-size:21px;font-weight:800}.stTextInput input{background:#fff!important;color:#111827!important;border:1px solid #d1d5db!important;border-radius:10px!important;min-height:44px}.stButton>button,.stDownloadButton>button{width:100%;min-height:44px;border-radius:10px!important;border:1px solid #2563eb!important;background:#2563eb!important;color:#fff!important;font-weight:700!important}button[kind="secondary"]{background:#fff!important;color:#2563eb!important}.result-card{background:#fff;border:1px solid #e5e7eb;border-radius:13px;padding:14px 16px;margin:8px 0}.result-name{font-size:15px;font-weight:800}.result-status{font-size:13px;font-weight:700;margin-top:2px}.status-badge{display:inline-block;padding:4px 10px;border-radius:999px;font-size:12px;font-weight:700}.status-ok{background:#ecfdf5;color:#15803d}.status-warn{background:#fffbeb;color:#b45309}.status-error{background:#fef2f2;color:#b91c1c}.status-info{background:#eff6ff;color:#1d4ed8}[data-testid="stFileUploaderDropzone"]{background:#f8fafc!important;border:1.5px dashed #cbd5e1!important;border-radius:14px!important}
</style>
""", unsafe_allow_html=True)

for key, default in [("logged_in",False),("is_admin",False),("office",None),("active_job_id",None)]:
    if key not in st.session_state: st.session_state[key] = default


# ==================== Login ====================
if not st.session_state.logged_in and not st.session_state.is_admin:
    st.markdown('<div style="text-align:center;margin:42px 0 24px;"><div style="font-size:48px;">✨</div><div style="font-size:31px;font-weight:800;color:#111827;">Aivora</div><div style="font-size:14px;color:#6b7280;">Your Smarter Support for Every Student\'s Application</div></div>', unsafe_allow_html=True)
    tab_login, tab_register, tab_admin = st.tabs(["تسجيل الدخول", "تسجيل مكتب جديد", "Admin"])
    with tab_login:
        name = st.text_input("اسم المكتب", key="login_office")
        password = st.text_input("الباسورد", type="password", key="login_password")
        if st.button("دخول", key="login_btn"):
            ok, result = check_login(name, password)
            if ok:
                st.session_state.logged_in = True
                st.session_state.office = result
                st.rerun()
            else: st.error(result)
    with tab_register:
        rn = st.text_input("اسم المكتب", key="reg_name")
        re_ = st.text_input("الإيميل", key="reg_email")
        rp = st.text_input("الباسورد", type="password", key="reg_password")
        if st.button("إنشاء الحساب", key="register_btn"):
            ok, msg = register_office(rn, re_, rp)
            (st.success if ok else st.error)(msg)
    with tab_admin:
        au = st.text_input("Admin username", key="admin_u")
        ap = st.text_input("Admin password", type="password", key="admin_p")
        if st.button("دخول الإدارة", key="admin_login"):
            if au == ADMIN_USERNAME and ap == ADMIN_PASSWORD:
                st.session_state.is_admin = True
                st.rerun()
            else: st.error("بيانات الإدارة غير صحيحة.")
    st.stop()


# ==================== Admin ====================
if st.session_state.is_admin:
    st.markdown('<div class="topbar"><div class="brand"><div class="brand-icon">✨</div><div><div class="brand-title">Aivora</div><div class="brand-sub">Administration</div></div></div></div>', unsafe_allow_html=True)
    st.markdown("## طلبات التسجيل")
    pending = get_pending_accounts()
    if not pending: st.info("مفيش طلبات تسجيل معلقة حاليًا.")
    for office in pending:
        c1,c2,c3,c4 = st.columns([2,2.2,1,1])
        with c1: st.write(f"**{office['name']}**")
        with c2: st.write(office.get('email',''))
        with c3:
            if st.button("قبول", key=f"approve_{office['id']}"):
                set_office_status(office['id'],"approved"); log_activity(office['id'],"تمت الموافقة على حساب المكتب"); st.rerun()
        with c4:
            if st.button("رفض", key=f"reject_{office['id']}"):
                set_office_status(office['id'],"rejected"); log_activity(office['id'],"تم رفض حساب المكتب"); st.rerun()
    st.divider()
    if st.button("تسجيل خروج الإدارة"):
        st.session_state.is_admin=False; st.rerun()
    st.stop()


# ==================== Office dashboard ====================
office = st.session_state.office
if not office:
    st.session_state.logged_in = False
    st.rerun()
office_id = office["id"]

st.markdown(f'<div class="topbar"><div class="brand"><div class="brand-icon">✨</div><div><div class="brand-title">Aivora</div><div class="brand-sub">{office["name"]}</div></div></div></div>', unsafe_allow_html=True)

if st.button("تسجيل الخروج", key="logout"):
    st.session_state.logged_in=False; st.session_state.office=None; st.session_state.active_job_id=None; st.rerun()

st.markdown('<div class="hero"><div class="hero-kicker">Aivora Agent</div><div class="hero-title">تابع طلبات طلابك <strong>بشكل أذكى</strong></div><div class="hero-desc">ارفع ملف الطلاب أو اربط Google Sheet، وخلي الـ Worker يتولى فحص الحالات.</div></div>', unsafe_allow_html=True)

students = get_students(office_id)
counts = {
    "total": len(students),
    "final": sum(str(s.get("application_status") or "") in FINAL_STATUSES for s in students),
    "pending": sum(str(s.get("application_status") or "") not in FINAL_STATUSES for s in students),
}
c1,c2,c3 = st.columns(3)
for col,label,value in [(c1,"إجمالي الطلاب",counts["total"]),(c2,"حالات نهائية",counts["final"]),(c3,"محتاج تحديث",counts["pending"])]:
    with col: st.markdown(f'<div class="stat-card"><div class="stat-label">{label}</div><div class="stat-value">{value}</div></div>',unsafe_allow_html=True)

st.markdown("## تحديث الطلبات")
source_mode = st.radio("مصدر بيانات الطلاب", ["رفع Excel", "Google Sheet"], horizontal=True)
file_bytes, source_url, file_name = None, None, "students.xlsx"
if source_mode == "رفع Excel":
    uploaded = st.file_uploader("اختاري ملف الطلاب", type=["xlsx"])
    if uploaded:
        file_bytes, file_name = uploaded.getvalue(), uploaded.name
else:
    source_url = st.text_input("رابط Google Sheet", placeholder="https://docs.google.com/spreadsheets/d/...")

if st.button("▶ ابدأ التحديث", key="start_update", disabled=(file_bytes is None and not source_url)):
    try:
        with st.spinner("بنجهز بيانات الطلاب في Supabase..."):
            source_type = "xlsx" if file_bytes is not None else "gsheet"
            source, count = import_students(office_id, source_type, file_name if file_bytes is not None else "Google Sheet", file_bytes=file_bytes, source_url=source_url)
            job = create_job(office_id, source, file_name)
            st.session_state.active_job_id = job["id"]
            log_activity(office_id, "إنشاء مهمة تحديث حالات", file_name, {"job_id": job["id"], "students": count}, data_source_id=source["id"])
        st.success(f"تم تجهيز {count} طالب. الـ Worker بدأ يستلم المهمة.")
    except Exception as exc:
        st.error(str(exc))

if st.session_state.active_job_id:
    job = get_job(st.session_state.active_job_id)
    if job:
        status = job.get("status", "pending")
        st.markdown(f"### حالة المهمة: `{status}`")
        progress = get_job_progress(job["id"])
        if progress:
            latest = progress[0]
            total = int(latest.get("total") or 0)
            current = int(latest.get("student_index") or 0)
            st.progress(min(current / total, 1.0) if total else 0.0)
            st.write(f"{current} / {total} — {latest.get('student_name','')} — {latest.get('status','')}")
        if status == "processing":
            time.sleep(2)
            st.rerun()
        elif status in {"done","failed"}:
            (st.success if status == "done" else st.error)("اكتملت المهمة." if status == "done" else (job.get("error") or "المهمة فشلت."))
            if st.button("إخفاء المهمة", key="hide_job"):
                st.session_state.active_job_id=None; st.rerun()

st.markdown("## البحث عن طالب")
search = st.text_input("اكتبي اسم الطالب", key="student_search")
if search:
    found = get_students(office_id, search)
    if not found: st.info("مش لاقي طالب بالاسم ده.")
    for student in found:
        status = student.get("application_status") or "لم يتم الفحص بعد"
        st.markdown(f'<div class="result-card"><div class="result-name">{student["student_name"]}</div><div class="result-status"><span class="status-badge {status_class(status)}">{status}</span></div></div>', unsafe_allow_html=True)

st.markdown("## حالات الطلاب")
if students:
    df = pd.DataFrame([{"الطالب":s["student_name"],"الحالة":s.get("application_status") or "لم يتم الفحص بعد","آخر تحديث":s.get("status_updated_at") or ""} for s in students])
    st.dataframe(df,use_container_width=True,hide_index=True)
else:
    st.info("لسه مفيش بيانات طلاب للمكتب ده.")

log_activity(office_id, "فتح لوحة المكتب")
