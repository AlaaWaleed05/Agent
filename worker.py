"""
worker_selenium.py
بديل worker.py — بيشتغل على جهازك المحلي (كروم ظاهر قدامك) بدل ما يعمل تسجيل
الدخول بـ requests. نفس آلية التنسيق مع Google Sheets (jobs / job_progress /
results) اللي app.py و auto_update.py معتمدين عليها من غير ما تحتاجي تعدّلي
فيهم حاجة — هما بيعرفوا إن فيه "Worker شغال" لما مهمة تتحول من pending
لـ processing، وده بيحصل هنا برضه بالظبط زي worker.py الأصلي.

الفرق الجوهري: بدل api_login / get_status اللي بيستخدموا requests.post على
الـ API مباشرة، هنا بنستخدم Selenium (كروم حقيقي ظاهر) وبنعمل login/فحص
حالة/logout زي إنسان عادي بيستخدم الموقع، بالظبط زي سكريبت "Agent - ادرس في
مصر" الأصلي اللي كان شغال كويس زمان.

قواعد التعامل مع الأخطاء:
- لو اللوجين فشل لأسباب طبيعية (إيميل/باسورد غلط) → حالة الطالب "فشل تسجيل
  الدخول" وبس، ده مش خطأ فني، بنكمل عادي من غير أي إيميل.
- لو حصل خطأ فني أثناء المعالجة (عنصر في الصفحة مش موجود، تايم آوت، الصفحة
  اتغيرت...) → بيتسجل للطالب ده "خطأ فني في الفحص" وبنكمل اللي بعده. الخطأ
  ده بيتجمع في قايمة، ولو حصل 5 أخطاء فنية ورا بعض على التوالي، بنعتبر إن
  فيه مشكلة نظامية (الموقع غيّر شكله مثلاً) فبنوقف المهمة على طول ونبعت
  إيميل عاجل، بدل ما نضيع وقت ولوجينات على الفاضي. في الآخر (سواء وقفت
  بدري أو خلصت عادي) لو فيه أي أخطاء فنية بنبعت إيميل ملخص بيهم لصاحب
  DEVELOPER_EMAIL.

طريقة الشغل:
    pip install gspread google-auth google-api-python-client openpyxl selenium
    export GCP_SERVICE_ACCOUNT_JSON='...محتوى ملف service account json...'
    export SENDER_EMAIL='...'            # (اختياري - لإرسال إيميلات الأخطاء)
    export SENDER_APP_PASSWORD='...'     # (اختياري)
    export DEVELOPER_EMAIL='...'         # (اختياري - هيستقبل إيميلات الأخطاء الفنية)
    python worker_selenium.py

أو حطي ملف service_account.json جنب السكريبت وشغليه من غير متغير بيئة.
محتاجة كمان Google Chrome مثبت على الجهاز (Selenium بيتحكم فيه مباشرة).
"""

import os
import io
import json
import time
import random
import socket
import uuid
import smtplib
from email.mime.text import MIMEText
from datetime import datetime

import openpyxl
import gspread
from google.oauth2.service_account import Credentials

from selenium import webdriver
from selenium.common.exceptions import WebDriverException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

# ==================== إعدادات ====================
SHEET_ID = "1BlFdtY-7ZIF1y2GwVosxlG9r7nK5xqYeW6yiIjPI_9U"
DRIVE_FOLDER_ID = "12L_qSHBnW4-tfQZRteynInWNBAML016f"

SITE_URL = "https://admission.study-in-egypt.gov.eg"
LOGIN_URL = f"{SITE_URL}/login"
INBOX_URL = f"{SITE_URL}/inbox"

WAIT_TIME = 20  # ثواني - أقصى انتظار لعنصر في الصفحة

# --- إعدادات موفّرة للـ Quota (نفس أرقام worker.py) ---
JOB_POLL_INTERVAL_SECONDS = 25
RESULTS_FLUSH_EVERY = 5
STUDENT_DELAY_MIN, STUDENT_DELAY_MAX = 5, 10  # فاصل إضافي بين طالب والتاني

# الحالات اللي لو الطالب وصلها، ما بنعيدش فحصه تاني
FINAL_STATUSES = {
    "مقبول نهائي",
    "قبول نهائي",
    "تم الرفض",
    "مرفوض نهائيًا",
    "مرفوض نهائيا",
    "مرفوض",
    "خالص",
}

TECH_FAILURE_STATUS = "خطأ فني في الفحص"
CONSECUTIVE_TECH_FAILURE_LIMIT = 5  # لو حصل كذا خطأ فني ورا بعض، نوقف المهمة

# ID فريد لكل نسخة Worker شغالة - عشان لو شغّلتي أكتر من Worker مع بعض، كل
# واحد يقدر "يمسك" مهمة من غير ما يتلخبط مع التاني
WORKER_ID = f"{socket.gethostname()}-{uuid.uuid4().hex[:6]}"

# لما Worker يلاقي مهمة pending، بيكتب الـ WORKER_ID بتاعه في عمود claimed_by
# ويستنى فترة عشوائية (عشان لو Worker تاني بيحاول يمسك نفس المهمة في نفس
# اللحظة تقريبًا يبان ده)، وبعدين يتأكد إن الـ ID المكتوب لسه بتاعه هو قبل ما
# يبدأ يشتغل عليها فعليًا.
CLAIM_JITTER_MIN, CLAIM_JITTER_MAX = 1.5, 3.5

# إعادة المحاولة عند 429 (Quota exceeded) على Google Sheets API
API_RETRY_MAX = 5
API_RETRY_BASE_WAIT = 20  # ثواني


# ==================== Retry wrapper لأي نداء Google Sheets API ====================

def with_retry(fn, *args, **kwargs):
    last_err = None
    for attempt in range(1, API_RETRY_MAX + 1):
        try:
            return fn(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            last_err = e
            msg = str(e)
            if "429" in msg or "Quota exceeded" in msg or "RESOURCE_EXHAUSTED" in msg:
                wait = API_RETRY_BASE_WAIT * attempt
                print(f"⚠️ Quota exceeded (attempt {attempt}/{API_RETRY_MAX}) — waiting {wait} seconds...")
                time.sleep(wait)
                continue
            raise
    raise last_err


# ==================== Credentials / Session (اتصال واحد يتكرر استخدامه) ====================

def get_creds():
    creds_json = os.environ.get("GCP_SERVICE_ACCOUNT_JSON")
    if creds_json:
        creds_dict = json.loads(creds_json)
    else:
        local_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "service_account.json")
        with open(local_path, "r", encoding="utf-8") as f:
            creds_dict = json.load(f)
    return Credentials.from_service_account_info(
        creds_dict,
        scopes=[
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
        ],
    )


class SheetSession:
    """اتصال واحد بالشيت الرئيسي، بيتفتح مرة واحدة، وبيكاش worksheets اللي
    بنستخدمها كتير (jobs/progress/results) عشان ما نفتحش اتصال جديد ونعمل
    lookup للتبويب في كل مرة."""

    def __init__(self):
        self.creds = get_creds()
        self.client = gspread.authorize(self.creds)
        self.spreadsheet = with_retry(self.client.open_by_key, SHEET_ID)
        self._ws_cache = {}
        self._office_ws_cache = {}

    def _ensure_ws(self, name, headers=None, rows=2000, cols=10):
        if name in self._ws_cache:
            return self._ws_cache[name]
        try:
            ws = with_retry(self.spreadsheet.worksheet, name)
        except Exception:
            ws = with_retry(self.spreadsheet.add_worksheet, name, rows, cols)
            if headers:
                with_retry(ws.append_row, headers)
        self._ws_cache[name] = ws
        return ws

    def jobs_ws(self):
        return self._ensure_ws(
            "jobs",
            headers=["job_id", "اسم المكتب", "نوع المصدر", "مرجع المصدر",
                     "اسم الملف", "الحالة", "تاريخ الإنشاء", "final_drive_file_id", "خطأ"],
            rows=500, cols=9,
        )

    def progress_ws(self):
        return self._ensure_ws(
            "job_progress",
            headers=["job_id", "index", "total", "اسم الطالب", "الحالة"],
            rows=5000, cols=5,
        )

    def results_ws(self):
        return self._ensure_ws(
            "results",
            headers=["اسم المكتب", "اسم الطالب", "الحالة", "تاريخ التحديث"],
            rows=5000, cols=4,
        )

    def accounts_ws(self):
        """شيت الحسابات (اسم المكتب/الإيميل/الباسورد/الحالة...) - نفس الشيت
        اللي app.py بيستخدمه لتسجيل الدخول. مبنستخدمهوش هنا غير عشان نجيب
        إيميل المكتب لبعت إيميلات تغيّر الحالة."""
        return self._ensure_ws(
            "accounts",
            headers=["اسم المكتب", "الإيميل", "الباسورد", "الحالة", "تاريخ التسجيل"],
            rows=1000, cols=5,
        )

    def claims_ws(self):
        """سجل إضافة بس (append-only) لمحاولات حجز المهام - كل سطر هنا بيتضاف
        وميتمسحش ولا يتكتب فوقه أبدًا، عشان نضمن ترتيب حقيقي وموثوق لمين
        حجز إيه أول، من غير أي لبس ممكن يحصل مع خلية بتتغيّر."""
        return self._ensure_ws(
            "job_claims",
            headers=["job_id", "worker_id", "وقت الحجز"],
            rows=5000, cols=3,
        )

    def office_gsheet_ws(self, link):
        sheet_id = extract_sheet_id(link)
        if not sheet_id:
            return None, "الرابط غير صحيح"
        cache_key = link
        if cache_key in self._office_ws_cache:
            return self._office_ws_cache[cache_key], None
        try:
            spreadsheet = with_retry(self.client.open_by_key, sheet_id)
            ws = get_target_worksheet(spreadsheet, link)
            self._office_ws_cache[cache_key] = ws
            return ws, None
        except Exception as e:
            return None, str(e)


# ==================== Sheets helpers ====================

def claim_next_pending_job(session: SheetSession):
    """بتدوّر على أول مهمة pending وتحاول 'تحجزها' عن طريق سجل إضافة بس
    (job_claims) - كل Worker بيسجل محاولة حجزه كسطر جديد (append)، والسطر ده
    مبيتمسحش ولا يتكتب فوقه أبدًا. القاعدة بسيطة وموثوقة 100%: أول سطر حجز
    وصل فعليًا للشيت لمهمة معينة هو اللي بيفوز بيها - وده Google Sheets نفسه
    بيحدده بشكل نهائي (بترتيب وصول الطلبات لسيرفره)، فمفيش أي احتمال لبس زي
    ما كان ممكن يحصل مع خلية بتتغيّر (last write wins)."""
    jobs_ws = session.jobs_ws()
    claims_ws = session.claims_ws()

    headers = with_retry(jobs_ws.row_values, 1)
    status_col = headers.index("الحالة") + 1

    job_records = with_retry(jobs_ws.get_all_records)
    pending_jobs = [(i, r) for i, r in enumerate(job_records, start=2) if str(r.get("الحالة", "")) == "pending"]

    for row_num, r in pending_jobs:
        job_id = str(r.get("job_id", "")).strip()
        if not job_id:
            continue

        # سجّلي محاولة حجزك - append بس، من غير أي كتابة فوق حاجة موجودة
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")
        with_retry(claims_ws.append_row, [job_id, WORKER_ID, now_str])

        # هامش أمان بسيط عشان أي تأخير طفيف في وصول الطلبات لسيرفر Sheets
        time.sleep(random.uniform(CLAIM_JITTER_MIN, CLAIM_JITTER_MAX))

        # اقرأي كل سجلات الحجز الخاصة بنفس المهمة دي بترتيب وصولها الفعلي
        all_claims = with_retry(claims_ws.get_all_records)
        claims_for_job = [c for c in all_claims if str(c.get("job_id", "")).strip() == job_id]

        if not claims_for_job or str(claims_for_job[0].get("worker_id", "")) != WORKER_ID:
            continue  # حد سبقك حجزها فعلاً - دوّري على مهمة تانية

        # فزتي بالحجز - اتأكدي كمان إن حالتها لسه pending فعلاً قبل ما تاخديها
        current_status = with_retry(jobs_ws.cell, row_num, status_col).value
        if current_status != "pending":
            continue

        r["_row"] = row_num
        return r

    return None


def get_next_pending_job(session: SheetSession):
    """نسخة قديمة (من غير حجز) - متسيبة هنا للرجوع ليها لو حبيتي، بس
    الاستخدام الافتراضي دلوقتي هو claim_next_pending_job الأقوى."""
    records = with_retry(session.jobs_ws().get_all_records)
    for i, r in enumerate(records, start=2):
        if str(r.get("الحالة", "")) == "pending":
            r["_row"] = i
            return r
    return None


def set_job_status(session: SheetSession, row_num, status, error="", final_drive_file_id=None):
    ws = session.jobs_ws()
    headers = with_retry(ws.row_values, 1)
    status_col = headers.index("الحالة") + 1
    with_retry(ws.update_cell, row_num, status_col, status)
    if error:
        error_col = headers.index("خطأ") + 1
        with_retry(ws.update_cell, row_num, error_col, str(error)[:400])
    if final_drive_file_id:
        fid_col = headers.index("final_drive_file_id") + 1
        with_retry(ws.update_cell, row_num, fid_col, final_drive_file_id)


def append_progress_batch(session: SheetSession, rows):
    if not rows:
        return
    with_retry(session.progress_ws().append_rows, rows, value_input_option="RAW")


def flush_results_to_sheet(session: SheetSession, office, results_so_far):
    try:
        ws = session.results_ws()
        all_values = with_retry(ws.get_all_values)
        headers = all_values[0] if all_values else ["اسم المكتب", "اسم الطالب", "الحالة", "تاريخ التحديث"]
        target = str(office).strip()
        kept_rows = [row for row in all_values[1:] if row and str(row[0]).strip() != target]
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        new_rows = [[target, r.get("name", ""), r.get("status", ""), now] for r in results_so_far]
        final_data = [headers] + kept_rows + new_rows
        with_retry(ws.clear)
        with_retry(ws.update, final_data)
    except Exception as e:
        print(f"Error saving results: {e}")


def log_to_sheet(session: SheetSession, office, action, filename=""):
    try:
        sheet = with_retry(session.client.open_by_key, SHEET_ID).sheet1
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        first_cell = with_retry(sheet.cell, 1, 1).value
        if sheet.row_count == 0 or first_cell != "التاريخ":
            with_retry(sheet.insert_row, ["التاريخ", "اسم المكتب", "العملية", "اسم الملف"], 1)
        with_retry(sheet.append_row, [now, office, action, filename])
    except Exception as e:
        print(f"Error logging action: {e}")


# ==================== Excel / GSheet helpers ====================

def extract_sheet_id(link):
    import re
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", link)
    return match.group(1) if match else None


def extract_gid(link):
    import re
    match = re.search(r"[?#&]gid=(\d+)", link)
    return int(match.group(1)) if match else None


def get_target_worksheet(spreadsheet, link):
    gid = extract_gid(link)
    if gid is not None:
        try:
            for ws in spreadsheet.worksheets():
                if ws.id == gid:
                    return ws
        except Exception:
            pass
    return spreadsheet.sheet1


def find_excel_columns(ws):
    cols = {"name": None, "email": None, "password": None, "status": None}
    header_row_num = None
    header_len = 0
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), start=1):
        row_values = [str(c).strip() if c else "" for c in row]
        if any("يميل" in v or "mail" in v.lower() for v in row_values):
            header_row_num = row_idx
            header_len = len(row_values)
            for i, cell in enumerate(row_values):
                if "حالة" in cell and "الجديدة" in cell:
                    cols["status"] = i
                    break
            for i, cell in enumerate(row_values):
                cell_lower = cell.lower()
                if any(k in cell for k in ["اسم", "الإسم", "الاسم"]) or "name" in cell_lower:
                    cols["name"] = i
                elif any(k in cell for k in ["يميل", "بريد"]) or "mail" in cell_lower:
                    cols["email"] = i
                elif any(k in cell for k in ["باسورد", "كلمة المرور", "password", "pass"]) or "pass" in cell_lower:
                    cols["password"] = i
            break
    if header_row_num is None:
        raise Exception("مش لاقي هيدر الإكسيل!")
    if cols["email"] is None:
        raise Exception("مش لاقي عمود الإيميل!")
    if cols["password"] is None:
        raise Exception("مش لاقي عمود الباسورد!")
    if cols["status"] is None:
        new_col_index = header_len
        ws.cell(row=header_row_num, column=new_col_index + 1, value="حالة الطلب الجديدة")
        cols["status"] = new_col_index
    return cols, header_row_num


def read_gsheet_as_workbook(session: SheetSession, link):
    ws, err = session.office_gsheet_ws(link)
    if err or not ws:
        return None, err or "تعذر الوصول للشيت"
    data = with_retry(ws.get_all_values)
    if not data:
        return None, "الشيت فاضي!"
    wb = openpyxl.Workbook()
    wsheet = wb.active
    for row in data:
        wsheet.append(row)
    return wb, None


def ensure_status_header_in_office_gsheet(session: SheetSession, link, header_row_num, status_col_index,
                                           header_name="حالة الطلب الجديدة"):
    ws, err = session.office_gsheet_ws(link)
    if err or not ws:
        return False
    col_num = status_col_index + 1
    current = with_retry(ws.cell, header_row_num, col_num).value
    if str(current or "").strip() != header_name:
        with_retry(ws.update_cell, header_row_num, col_num, header_name)
    return True


def update_status_cell_in_office_gsheet(session: SheetSession, link, row_num, status_col_index, value):
    ws, err = session.office_gsheet_ws(link)
    if err or not ws:
        return False, err or "تعذر الوصول للشيت"
    with_retry(ws.update_cell, row_num, status_col_index + 1, value)
    return True, None


def get_previous_results_for_office(session: SheetSession, office):
    try:
        ws = session.results_ws()
        all_values = with_retry(ws.get_all_values)
        if len(all_values) <= 1:
            return {}
        target = str(office).strip()
        previous = {}
        for row in all_values[1:]:
            if len(row) < 3:
                continue
            if str(row[0]).strip() == target:
                previous[str(row[1]).strip()] = str(row[2]).strip()
        return previous
    except Exception as e:
        print(f"Could not fetch previous results for office {office}: {e}")
        return {}


def get_office_email(session: SheetSession, office):
    """بترجع إيميل المكتب المسجل في شيت accounts (أو None لو مش لاقياه)."""
    try:
        ws = session.accounts_ws()
        records = with_retry(ws.get_all_records)
        target = str(office).strip()
        for r in records:
            if str(r.get("اسم المكتب", "")).strip() == target:
                email = str(r.get("الإيميل", "") or "").strip()
                return email or None
        return None
    except Exception as e:
        print(f"Could not fetch office email for {office}: {e}")
        return None


def compute_status_changes(previous_results, new_results):
    """بترجع قايمة الطلاب اللي حالتهم اتغيّرت (أو ظهروا لأول مرة) بمقارنة
    النتائج القديمة بالجديدة."""
    changes = []
    for r in new_results:
        name = str(r.get("name", "")).strip()
        new_status = str(r.get("status", "")).strip()
        if not name:
            continue
        old_status = previous_results.get(name)
        if old_status is None or old_status != new_status:
            changes.append({"name": name, "new_status": new_status, "old_status": old_status})
    return changes


def build_status_change_email(office, changes):
    lines = [f"تحديث حالات الطلاب - {office}", ""]
    shown = changes[:200]
    for idx, c in enumerate(shown, start=1):
        lines.append(f"{idx}. {c['name']} - {c['new_status']}")
    if len(changes) > 200:
        lines.append("")
        lines.append(f"...و{len(changes) - 200} طالب/ة تاني اتغيرت حالتهم")
    lines.append("")
    lines.append(f"إجمالي التغييرات: {len(changes)}")
    return "\n".join(lines)


def notify_office_status_changes(session: SheetSession, office, previous_results, processed_results):
    """لو فيه طلاب اتغيرت حالتهم (أو ظهروا لأول مرة) في التحديث ده، بتبعت
    إيميل واحد للمكتب فيهم بس. لو ده أول تحديث للمكتب ده على الإطلاق (مفيش
    نتائج سابقة خالص)، مبنبعتش إيميل - عشان منغرقوش المكتب بإيميل فيه كل
    الطلاب كأنهم 'اتغيروا'."""
    if not previous_results:
        print("First-ever check for this office - skipping change-notification email.")
        return
    changes = compute_status_changes(previous_results, processed_results)
    if not changes:
        return
    office_email = get_office_email(session, office)
    if not office_email:
        print(f"No email on file for office {office} - skipping change-notification email.")
        return
    subject = f"تحديث حالات الطلاب - {office}"
    body = build_status_change_email(office, changes)
    send_email_notification(office_email, subject, body)
    print(f"✉️ Sent status-change email to {office} ({len(changes)} change(s)).")


def download_drive_file(creds, file_id):
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload

    service = build("drive", "v3", credentials=creds)
    request = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return buf.read()


def upload_final_file(creds, file_bytes, filename, office):
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseUpload

    service = build("drive", "v3", credentials=creds)
    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    drive_filename = f"{office}_{now}_{filename}"
    file_metadata = {"name": drive_filename, "parents": [DRIVE_FOLDER_ID]}
    media = MediaIoBaseUpload(
        io.BytesIO(file_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    created = service.files().create(body=file_metadata, media_body=media, fields="id", supportsAllDrives=True).execute()
    return created.get("id")


# ==================== إيميل تنبيهات الأخطاء الفنية ====================

def send_email_notification(to_email, subject, body):
    sender_email = os.environ.get("SENDER_EMAIL")
    sender_password = os.environ.get("SENDER_APP_PASSWORD")

    if not sender_email or not sender_password:
        print("ℹ️ Email settings incomplete (SENDER_EMAIL/SENDER_APP_PASSWORD) - skipping send.")
        return
    if not to_email:
        print("ℹ️ No recipient email set - skipping send.")
        return
    try:
        msg = MIMEText(body, "plain", "utf-8")
        msg["Subject"] = subject
        msg["From"] = sender_email
        msg["To"] = to_email
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, [to_email], msg.as_string())
        print(f"✉️ Email sent to {to_email}")
    except Exception as e:
        print(f"❌ Error sending email: {e}")


def notify_developer_tech_failures(office, job_id, failures, stopped_early=False):
    """بتبعت إيميل ملخص للأخطاء الفنية اللي حصلت أثناء مهمة معينة (لو فيه)."""
    developer_email = os.environ.get("DEVELOPER_EMAIL")
    if not developer_email or not failures:
        return
    title = "⚠️ أخطاء فنية أثناء تحديث حالات الطلاب"
    if stopped_early:
        title = "🛑 " + title + " (تم إيقاف المهمة مبكرًا)"
    lines = [f"{title}", f"المكتب: {office}", f"job_id: {job_id}", ""]
    if stopped_early:
        lines.append(
            f"تم إيقاف المهمة تلقائيًا بعد {CONSECUTIVE_TECH_FAILURE_LIMIT} أخطاء فنية متتالية — "
            "على الأرجح الموقع غيّر شكل صفحة اللوجين/الطلبات، يستحسن تتأكدي بنفسك على المتصفح."
        )
        lines.append("")
    lines.append(f"عدد الأخطاء الفنية: {len(failures)}")
    lines.append("")
    for i, f in enumerate(failures[:100], start=1):
        lines.append(f"{i}. {f['name']}: {f['error']}")
    if len(failures) > 100:
        lines.append(f"...و{len(failures) - 100} خطأ تاني")
    send_email_notification(developer_email, title + f" - {office}", "\n".join(lines))


# ==================== Selenium: أدوات عامة ====================

def human_type(element, text):
    element.clear()
    time.sleep(1)
    for char in text:
        element.send_keys(char)
        time.sleep(random.uniform(0.15, 0.4))


def slow_wait(seconds, msg=""):
    if msg:
        print(f"    ⏳ {msg} ({seconds}s)...")
    time.sleep(seconds)


def setup_browser():
    options = Options()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument("--start-maximized")
    # مفيش user-agent مكتوب يدوي هنا عن قصد - سيبنا الكروم يبعت الـ user-agent
    # الحقيقي بتاعه (بإصداره الفعلي المثبت عندك)، عشان ميبقاش فيه تعارض بين
    # رقم الإصدار في الهيدر ورقم الإصدار الفعلي للمتصفح المتصل.
    driver = webdriver.Chrome(options=options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver


def clear_session(driver):
    """بتمسح جلسة الطالب بالكامل: الكوكيز + localStorage + sessionStorage.
    مسح الكوكيز بس مش كافي مع مواقع SPA حديثة زي ده، اللي بتحفظ توكن الدخول
    في localStorage/sessionStorage - لو التوكن ده فضل موجود، الموقع بيفضل
    شايف الطالب 'داخل بالفعل' حتى لو مسحنا الكوكيز، وبيحوّل الطالب اللي بعده
    تلقائيًا بعيد عن صفحة اللوجين (فحقل الإيميل مايبقاش موجود ويطلع خطأ فني)."""
    try:
        driver.delete_all_cookies()
    except Exception:
        pass
    try:
        driver.execute_script("window.localStorage.clear(); window.sessionStorage.clear();")
    except Exception:
        pass


def selenium_login(driver, email, password):
    """بترجع (ok, is_technical_error, error_message).
    ok=True يعني اللوجين نجح.
    لو ok=False و is_technical_error=False → فشل لوجين طبيعي (إيميل/باسورد غلط).
    لو is_technical_error=True → في مشكلة في الصفحة نفسها (عنصر مش موجود...)."""
    try:
        driver.get(LOGIN_URL)
        # احتياط إضافي: نمسح أي جلسة/توكن قديم عالق من طالب سابق (حتى لو
        # افتراضًا كان المفروض اتمسح بالفعل) قبل ما نتأكد إننا فعلاً واقفين
        # في صفحة اللوجين الحقيقية، وبعدين نعيد تحميل الصفحة عشان أي تحويل
        # تلقائي (auto-redirect) بناءً على الجلسة القديمة يتلغي.
        clear_session(driver)
        driver.get(LOGIN_URL)
        wait = WebDriverWait(driver, WAIT_TIME)
        slow_wait(3, "Loading login page")

        email_field = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "input[type='email'], input[name='email']")
        ))
        email_field.click()
        slow_wait(3, "Waiting before email")
        human_type(email_field, email)

        password_field = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
        password_field.click()
        slow_wait(3, "Waiting before password")
        human_type(password_field, password)

        slow_wait(2, "Waiting before click")
        login_btn = driver.find_element(By.CSS_SELECTOR, "button[type='submit'], input[type='submit']")
        login_btn.click()

        # بدل ما نستنى وقت ثابت (5 ثواني) وناخد "لقطة واحدة" بعدها مباشرة،
        # بنفضل نتأكد كل ثانية لحد 15 ثانية - عشان لو التحويل من صفحة اللوجين
        # لصفحة الطلبات كان بطيء شوية (سيرفر مشغول، نت بطيء...)، منحكمش
        # "فشل" بدري قبل ما يكمل فعليًا. أول ما حقل الإيميل يختفي، نعتبرها
        # نجحت على طول من غير ما نستنى الـ 15 ثانية كاملة.
        login_success = False
        for _ in range(15):
            time.sleep(1)
            try:
                driver.find_element(By.CSS_SELECTOR, "input[type='email'], input[name='email']")
                still_has_email_field = True
            except Exception:
                still_has_email_field = False

            if not still_has_email_field or "login" not in driver.current_url:
                login_success = True
                break

        if not login_success:
            return False, False, "فشل تسجيل الدخول"
        return True, False, None

    except (TimeoutException, WebDriverException) as e:
        return False, True, f"خطأ فني في صفحة اللوجين: {e}"
    except Exception as e:
        return False, True, f"خطأ غير متوقع في اللوجين: {e}"


def selenium_go_to_inbox(driver):
    """بتضغط على القائمة وبعدين 'طلباتي' زي إنسان عادي بيتنقل، بتوقيت سريع.
    لو حصل أي مشكلة في الضغط (عنصر مش لاقيينه)، بترجع لصفحة الطلبات مباشرة
    كخطة بديلة بدل ما توقف تمامًا."""
    try:
        wait = WebDriverWait(driver, WAIT_TIME)
        try:
            menu_btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(text(), 'القائمة') or contains(@class,'menu')]")
            ))
            menu_btn.click()
            slow_wait(0.8, "Waiting for menu to open")
            my_apps = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(text(), 'طلباتي')]")
            ))
            my_apps.click()
            slow_wait(1, "Waiting for requests page")
        except Exception:
            driver.get(INBOX_URL)
            slow_wait(1)
        return True, None
    except (TimeoutException, WebDriverException) as e:
        return False, f"خطأ فني في الوصول لصفحة الطلبات: {e}"


def selenium_get_status(driver):
    """بترجع (status, is_technical_error, error_message)."""
    try:
        wait = WebDriverWait(driver, WAIT_TIME)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table")))
        slow_wait(2)

        headers = driver.find_elements(By.CSS_SELECTOR, "table thead th, table tr th")
        header_texts = [h.text.strip() for h in headers]

        status_index = None
        date_index = None
        for i, h in enumerate(header_texts):
            if h in ("حالة الطلب", "الحالة"):
                status_index = i
            if "آخر تعديل" in h:
                date_index = i
        if status_index is None:
            for i, h in enumerate(header_texts):
                if "حالة" in h and "اسم" not in h and "خدمة" not in h:
                    status_index = i
                    break

        rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        all_requests = []
        for row in rows:
            cells = row.find_elements(By.CSS_SELECTOR, "td")
            if not cells:
                continue
            status = cells[status_index].text.strip() if status_index is not None and status_index < len(cells) else ""
            if status:
                all_requests.append(status)

        if not all_requests:
            return "مفيش طلبات", False, None

        return all_requests[0], False, None

    except (TimeoutException, WebDriverException) as e:
        return "", True, f"خطأ فني في جلب حالة الطلب: {e}"
    except Exception as e:
        return "", True, f"خطأ غير متوقع في جلب حالة الطلب: {e}"


def selenium_logout(driver):
    """بتحاول تعمل logout عن طريق الواجهة (زي إنسان عادي)، وفي كل الأحوال -
    نجح الزرار ولا لأ - بتمسح كل الكوكيز في النهاية. ده مهم جدًا: لو حصل خطأ
    في تحديد 'هل الطالب داخل فعلاً ولا لأ' (زي ما حصل قبل كده)، مسح الكوكيز
    هنا بيضمن إن جلسة الطالب ده ما تفضلش عالقة وتلخبط كل الطلاب اللي بعده."""
    try:
        wait = WebDriverWait(driver, WAIT_TIME)
        slow_wait(2)
        user_menu = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "[class*='user'], [class*='profile'], [class*='avatar'], [class*='account']")
        ))
        user_menu.click()
        slow_wait(2)
        logout_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[contains(text(), 'تسجيل خروج') or contains(text(), 'خروج')]")
        ))
        logout_btn.click()
        slow_wait(2)
    except Exception:
        pass
    finally:
        clear_session(driver)


def restart_browser(driver):
    """بتقفل الكروم القديم (لو لسه شغال) وتفتح واحد جديد - بتتنادى لو الكروم
    وقع تمامًا (مش مجرد عنصر مش لاقيينه)."""
    try:
        driver.quit()
    except Exception:
        pass
    return setup_browser()


# ==================== Core: تنفيذ مهمة ====================

def process_job(session: SheetSession, job):
    job_id = job["job_id"]
    office = job["اسم المكتب"]
    source_type = job["نوع المصدر"]
    source_ref = job["مرجع المصدر"]
    filename = job.get("اسم الملف") or "students.xlsx"
    row_num = job["_row"]

    print(f"=== Starting job {job_id} for office {office} (Selenium) ===")
    set_job_status(session, row_num, "processing")

    try:
        if source_type == "gsheet":
            wb, err = read_gsheet_as_workbook(session, source_ref)
            if err:
                set_job_status(session, row_num, "failed", error=err)
                return
        else:  # drive_pending / drive
            file_bytes = download_drive_file(session.creds, source_ref)
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes))

        ws = wb.active
        cols, header_row_num = find_excel_columns(ws)
    except Exception as e:
        set_job_status(session, row_num, "failed", error=str(e))
        return

    previous_results = get_previous_results_for_office(session, office)

    if source_type == "gsheet":
        ensure_status_header_in_office_gsheet(session, source_ref, header_row_num, cols["status"])

    rows_data = list(ws.iter_rows(min_row=header_row_num + 1, values_only=False))
    all_valid_rows = [r for r in rows_data if r[cols["email"]].value and r[cols["password"]].value]

    pending_rows = []
    for r in all_valid_rows:
        name = str(r[cols["name"]].value or "").strip() if cols["name"] is not None else ""
        if previous_results.get(name) in FINAL_STATUSES:
            continue
        pending_rows.append(r)

    total = len(pending_rows)
    if total == 0:
        set_job_status(session, row_num, "failed", error="مفيش طلاب محتاجين تحديث (كلهم في حالة نهائية أو الملف فاضي)")
        return

    processing_order = pending_rows[:]
    random.shuffle(processing_order)

    processed_results = []
    progress_buffer = []
    tech_failures = []
    tech_retry_list = []  # الطلاب اللي طلعلهم "خطأ فني" - هنعيد محاولتهم في الآخر
    consecutive_tech_failures = 0
    stopped_early = False

    print("🖥️  Opening Chrome...")
    driver = setup_browser()

    try:
        for idx, row in enumerate(processing_order):
            email = str(row[cols["email"]].value).strip()
            password = str(row[cols["password"]].value).strip()
            name = row[cols["name"]].value if cols["name"] is not None else ""
            display_name = name or email
            row_num_in_sheet = row[cols["email"]].row

            print(f"\n👤 [{idx+1}/{total}] {display_name}")

            current_status = None
            is_tech_error = False
            error_msg = None
            browser_crashed = False
            login_confirmed_failed = False  # صح بس لما نتأكد إن اللوجين فشل فعلاً (مش دخل خالص)

            try:
                ok, is_tech_error, error_msg = selenium_login(driver, email, password)
                if not ok and not is_tech_error:
                    current_status = "فشل تسجيل الدخول"
                    login_confirmed_failed = True
                elif not ok and is_tech_error:
                    current_status = TECH_FAILURE_STATUS
                else:
                    ok2, err2 = selenium_go_to_inbox(driver)
                    if not ok2:
                        is_tech_error = True
                        error_msg = err2
                        current_status = TECH_FAILURE_STATUS
                    else:
                        status_text, is_tech_error, error_msg = selenium_get_status(driver)
                        if is_tech_error:
                            current_status = TECH_FAILURE_STATUS
                        else:
                            current_status = status_text
            except WebDriverException as e:
                # الكروم وقع تمامًا - نفتح واحد جديد ونعتبرها خطأ فني للطالب ده
                is_tech_error = True
                error_msg = f"الكروم اتقفل/وقع أثناء المعالجة: {e}"
                current_status = TECH_FAILURE_STATUS
                browser_crashed = True
            finally:
                # مهم جدًا: بنعمل logout/مسح جلسة بعد *كل* طالب بغض النظر عن
                # النتيجة (نجح الدخول، فشل، أو حصل خطأ فني) - عشان لو الفحص
                # بتاعنا غلط في تحديد "نجح ولا فشل" (زي ما حصل قبل كده)،
                # الجلسة القديمة ما تفضلش شغالة وتكسر كل الطلاب اللي بعده.
                if browser_crashed:
                    driver = restart_browser(driver)
                elif login_confirmed_failed:
                    # لوجين فشل بشكل مؤكد = مفيش جلسة أصلاً نعمل ليها logout.
                    # منحاولش نضغط على قائمة مستخدم مش موجودة (كان بياخد لغاية
                    # 20 ثانية فاضية لغاية ما يطلع timeout) - بس نمسح الجلسة
                    # كاملة (كوكيز + localStorage/sessionStorage) على طول احتياطًا.
                    clear_session(driver)
                else:
                    try:
                        selenium_logout(driver)
                    except Exception:
                        clear_session(driver)

            if is_tech_error:
                consecutive_tech_failures += 1
                tech_failures.append({"name": str(display_name), "error": error_msg or "خطأ غير معروف"})
                print(f"    ⚠️ Technical error: {error_msg}")
                if current_status == TECH_FAILURE_STATUS:
                    tech_retry_list.append({
                        "row": row,
                        "row_num_in_sheet": row_num_in_sheet,
                        "email": email,
                        "password": password,
                        "display_name": display_name,
                    })
            else:
                consecutive_tech_failures = 0
                print(f"    ✅ Status: {current_status}")

            if cols["status"] is not None:
                row[cols["status"]].value = current_status

            processed_results.append({"name": str(display_name), "status": str(current_status)})
            progress_buffer.append([job_id, idx + 1, total, str(display_name), str(current_status)])

            if source_type == "gsheet":
                update_status_cell_in_office_gsheet(session, source_ref, row_num_in_sheet, cols["status"], current_status)

            if len(progress_buffer) >= RESULTS_FLUSH_EVERY or idx == total - 1:
                append_progress_batch(session, progress_buffer)
                progress_buffer = []
                flush_results_to_sheet(session, office, processed_results)

        

            if idx < total - 1:
                slow_wait(random.uniform(STUDENT_DELAY_MIN, STUDENT_DELAY_MAX), "Pause before next student")

        # =================================================================
        # إعادة محاولة تلقائية لمين طلعله "خطأ فني في الفحص" بعد ما نخلص من
        # كل الطلاب - غير مؤكدين إن فيه مشكلة حقيقية في بياناته، ممكن تكون
        # مجرد مشكلة تقنية عابرة استحقت فرصة تانية (زي حالة الـ SSL/اتصال
        # المؤقتة في المسار التاني بتاع app.py بالظبط).
        # =================================================================
        if tech_retry_list:
            print(f"\n🔁 Retrying {len(tech_retry_list)} student(s) with technical errors...")
            for retry_item in tech_retry_list:
                r_row = retry_item["row"]
                r_email = retry_item["email"]
                r_password = retry_item["password"]
                r_display_name = retry_item["display_name"]
                r_row_num = retry_item["row_num_in_sheet"]

                print(f"\n🔁 Retry: {r_display_name}")
                slow_wait(random.uniform(STUDENT_DELAY_MIN, STUDENT_DELAY_MAX), "Pause before retry")

                retry_status = None
                retry_is_tech_error = False
                retry_error_msg = None
                retry_login_confirmed_failed = False
                retry_browser_crashed = False

                try:
                    ok, retry_is_tech_error, retry_error_msg = selenium_login(driver, r_email, r_password)
                    if not ok and not retry_is_tech_error:
                        retry_status = "فشل تسجيل الدخول"
                        retry_login_confirmed_failed = True
                    elif not ok and retry_is_tech_error:
                        retry_status = TECH_FAILURE_STATUS
                    else:
                        ok2, err2 = selenium_go_to_inbox(driver)
                        if not ok2:
                            retry_is_tech_error = True
                            retry_error_msg = err2
                            retry_status = TECH_FAILURE_STATUS
                        else:
                            status_text, retry_is_tech_error, retry_error_msg = selenium_get_status(driver)
                            retry_status = TECH_FAILURE_STATUS if retry_is_tech_error else status_text
                except WebDriverException as e:
                    retry_is_tech_error = True
                    retry_error_msg = f"الكروم اتقفل/وقع أثناء إعادة المحاولة: {e}"
                    retry_status = TECH_FAILURE_STATUS
                    retry_browser_crashed = True
                finally:
                    if retry_browser_crashed:
                        driver = restart_browser(driver)
                    elif retry_login_confirmed_failed:
                        clear_session(driver)
                    else:
                        try:
                            selenium_logout(driver)
                        except Exception:
                            clear_session(driver)

                if retry_is_tech_error:
                    print(f"    ⚠️ Still a technical error after retry: {retry_error_msg}")
                else:
                    print(f"    ✅ Status after retry: {retry_status}")

                if cols["status"] is not None:
                    r_row[cols["status"]].value = retry_status

                for pr in processed_results:
                    if pr["name"] == str(r_display_name):
                        pr["status"] = str(retry_status)
                        break

                if source_type == "gsheet":
                    update_status_cell_in_office_gsheet(session, source_ref, r_row_num, cols["status"], retry_status)

            # حدّثي شيت النتائج بعد كل إعادة المحاولات
            flush_results_to_sheet(session, office, processed_results)

            # لو حد اتظبط بعد إعادة المحاولة، اشيليه من قايمة الأخطاء اللي
            # هتتبعت للمطوّرة بالإيميل - مبقاش فيه مشكلة فعلية بقى
            tech_failures = [
                f for f in tech_failures
                if any(pr["name"] == f["name"] and pr["status"] == TECH_FAILURE_STATUS for pr in processed_results)
            ]

    finally:
        try:
            driver.quit()
        except Exception:
            pass

    # لو فيه طلاب اتغيرت حالتهم عن آخر مرة، ابعتي إيميل واحد للمكتب فيهم بس
    try:
        notify_office_status_changes(session, office, previous_results, processed_results)
    except Exception as e:
        print(f"Could not send status-change email: {e}")

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

    final_file_id = None
    try:
        final_file_id = upload_final_file(session.creds, out.getvalue(), filename, office)
    except Exception as e:
        print(f"⚠️ Warning: failed to upload final file to Drive: {e}")

    log_to_sheet(session, office, "اكتمل المعالجة عبر Worker (Selenium)", filename)

    if stopped_early:
        set_job_status(
            session, row_num, "failed",
            error=f"تم إيقاف المهمة بعد {CONSECUTIVE_TECH_FAILURE_LIMIT} أخطاء فنية متتالية",
            final_drive_file_id=final_file_id or "",
        )
    else:
        set_job_status(session, row_num, "done", final_drive_file_id=final_file_id or "")

    if tech_failures:
        notify_developer_tech_failures(office, job_id, tech_failures, stopped_early=stopped_early)

    print(f"=== Job {job_id} finished ({len(processed_results)}/{total} students"
          f"{', stopped early' if stopped_early else ''}) ===")


# ==================== Main loop ====================

def main():
    print(f"Worker (Selenium) running... ID: {WORKER_ID}")
    print(f"Checking for new jobs every ~{JOB_POLL_INTERVAL_SECONDS} seconds.")
    print("Chrome will open visibly as soon as a new job arrives.")
    session = SheetSession()

    while True:
        try:
            job = claim_next_pending_job(session)
            if job:
                process_job(session, job)
            else:
                # تشتيت بسيط حوالي فترة الفحص - عشان لو شغّلتي أكتر من Worker
                # مع بعض، ميبقوش بيفحصوا في نفس اللحظة بالظبط كل مرة
                time.sleep(random.uniform(JOB_POLL_INTERVAL_SECONDS * 0.7, JOB_POLL_INTERVAL_SECONDS * 1.3))

        except KeyboardInterrupt:
            print("Worker stopped manually.")
            break
        except gspread.exceptions.APIError as e:
            msg = str(e)
            if "429" in msg or "Quota exceeded" in msg:
                print("⚠️ Quota exceeded in main loop — waiting a full minute before continuing...")
                time.sleep(60)
            else:
                print(f"Unexpected API error: {e}")
                time.sleep(10)
        except Exception as e:
            print(f"Unexpected error in main loop: {e}")
            time.sleep(10)


if __name__ == "__main__":
    main()