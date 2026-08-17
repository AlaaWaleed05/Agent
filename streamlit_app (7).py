"""
ادرس في مصر - Agent
Streamlit App
"""

import streamlit as st
import gspread
import plotly.express as px
import pandas as pd
from google.oauth2.service_account import Credentials
from datetime import datetime
import requests
import openpyxl
import json
import io
import time
import random

# ==================== إعدادات ====================
ADMIN_USERNAME = "admin"   # ← اسم المسؤول
ADMIN_PASSWORD = "admin123"  # ← باسورد المسؤول

SHEET_ID = "1BlFdtY-7ZIF1y2GwVosxlG9r7nK5xqYeW6yiIjPI_9U"
DRIVE_FOLDER_ID = "12L_qSHBnW4-tfQZRteynInWNBAML016f"

# ==================== Google Sheets Logging ====================

def get_sheet():
    """اتصل بـ Google Sheets"""
    try:
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(
            creds_dict,
            scopes=["https://spreadsheets.google.com/feeds",
                    "https://www.googleapis.com/auth/drive"]
        )
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SHEET_ID).sheet1
        return sheet
    except Exception as e:
        print(f"خطأ في Google Sheets: {e}")
        return None

def log_to_sheet(office, action, filename=""):
    """بيسجل كل عملية في Google Sheets"""
    try:
        sheet = get_sheet()
        if sheet:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            # تأكد إن الهيدر موجود
        if sheet.row_count == 0 or sheet.cell(1,1).value != "التاريخ":
            sheet.insert_row(["التاريخ", "اسم المكتب", "العملية", "اسم الملف"], 1)
        sheet.append_row([now, office, action, filename])
    except Exception as e:
        print(f"خطأ في التسجيل: {e}")

def get_accounts_sheet():
    """جيب شيت الحسابات"""
    try:
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(
            creds_dict,
            scopes=["https://spreadsheets.google.com/feeds",
                    "https://www.googleapis.com/auth/drive"]
        )
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(SHEET_ID)
        try:
            sheet = spreadsheet.worksheet("accounts")
        except:
            sheet = spreadsheet.add_worksheet("accounts", 1000, 5)
            sheet.append_row(["اسم المكتب", "الإيميل", "الباسورد", "الحالة", "تاريخ التسجيل"])
        return sheet
    except Exception as e:
        return None

def register_office(office_name, email, password):
    """تسجيل مكتب جديد"""
    try:
        sheet = get_accounts_sheet()
        if not sheet:
            return False, "خطأ في الاتصال"
        
        # تأكد مش مسجل قبل كده
        records = sheet.get_all_records()
        for r in records:
            if r.get("اسم المكتب") == office_name or r.get("الإيميل") == email:
                return False, "اسم المكتب أو الإيميل مسجل قبل كده!"
        
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        sheet.append_row([office_name, email, password, "pending", now])
        return True, "تم التسجيل! في انتظار موافقة الإدارة."
    except Exception as e:
        return False, str(e)

def check_login(office_name, password):
    """تحقق من اللوجين"""
    try:
        sheet = get_accounts_sheet()
        if not sheet:
            return False, "خطأ في الاتصال"
        
        records = sheet.get_all_records()
        for r in records:
            if r.get("اسم المكتب") == office_name and r.get("الباسورد") == password:
                if r.get("الحالة") == "approved":
                    return True, "ok"
                elif r.get("الحالة") == "pending":
                    return False, "حسابك في انتظار موافقة الإدارة!"
                else:
                    return False, "حسابك موقوف!"
        return False, "اسم المكتب أو الباسورد غلط!"
    except Exception as e:
        return False, str(e)

def get_pending_accounts():
    """جيب الحسابات في انتظار الموافقة"""
    try:
        sheet = get_accounts_sheet()
        if not sheet:
            return []
        records = sheet.get_all_records()
        return [r for r in records if r.get("الحالة") == "pending"]
    except:
        return []

def approve_account(office_name):
    """وافقي على حساب"""
    try:
        sheet = get_accounts_sheet()
        records = sheet.get_all_records()
        for i, r in enumerate(records, start=2):
            if r.get("اسم المكتب") == office_name:
                sheet.update_cell(i, 4, "approved")
                return True
        return False
    except:
        return False

def reject_account(office_name):
    """ارفضي حساب"""
    try:
        sheet = get_accounts_sheet()
        records = sheet.get_all_records()
        for i, r in enumerate(records, start=2):
            if r.get("اسم المكتب") == office_name:
                sheet.update_cell(i, 4, "rejected")
                return True
        return False
    except:
        return False

def get_latest_file_from_drive(office):
    """بيجيب آخر إكسيل محدث للمكتب من Drive"""
    try:
        from googleapiclient.discovery import build
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(
            creds_dict,
            scopes=["https://www.googleapis.com/auth/drive"]
        )
        service = build("drive", "v3", credentials=creds)

        # اسم المكتب ممكن يكون فيه علامات تكسر الكويري (زي ' أو \) — نعمل escape ليها
        safe_office = str(office).replace("\\", "\\\\").replace("'", "\\'")

        # ابحث عن آخر ملف للمكتب ده في الفولدر (مع دعم Shared Drives لو الفولدر جواها)
        results = service.files().list(
            q=f"'{DRIVE_FOLDER_ID}' in parents and name contains '{safe_office}' and trashed=false",
            orderBy="createdTime desc",
            pageSize=1,
            fields="files(id, name, createdTime)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()

        files = results.get("files", [])
        if not files:
            # نتأكد الفولدر أصلاً متاح ومقدر نشوف اللي فيه، عشان نميز
            # بين "مفيش ملفات بالاسم ده" و"مش قادرين نوصل للفولدر خالص"
            try:
                probe = service.files().list(
                    q=f"'{DRIVE_FOLDER_ID}' in parents and trashed=false",
                    pageSize=5,
                    fields="files(id, name)",
                    supportsAllDrives=True,
                    includeItemsFromAllDrives=True,
                ).execute()
                existing = [f["name"] for f in probe.get("files", [])]
                if not existing:
                    return None, "الفولدر على Drive فاضي أو الـ service account مالوش صلاحية يشوف محتوياته — تأكدي إن الفولدر متشيير مع إيميل الـ service account."
                return None, f"مفيش ملفات باسم فيه '{office}'. أسماء ملفات موجودة فعلاً في الفولدر: {existing}"
            except Exception as probe_err:
                return None, f"مفيش ملفات محفوظة للمكتب ده، وكمان فشل فحص الفولدر: {probe_err}"

        # حمل الملف
        from googleapiclient.http import MediaIoBaseDownload
        import io as _io

        file_id = files[0]["id"]
        request = service.files().get_media(fileId=file_id)
        buf = _io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        buf.seek(0)
        return buf, None

    except Exception as e:
        return None, str(e)


def search_in_excel(file_bytes, name_query):
    """بيبحث في الإكسيل عن طالب بالاسم ويرجع آخر حالة"""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        ws = wb.active

        # لاقي الهيدر
        cols, header_row_num = find_excel_columns(ws)

        results = []
        for row in ws.iter_rows(min_row=header_row_num + 1, values_only=True):
            name   = str(row[cols["name"]]   if cols["name"]   is not None else "").strip()
            status = str(row[cols["status"]] if cols["status"] is not None else "").strip()

            if name_query.strip() in name:
                results.append({"name": name, "status": status})

        return results
    except Exception as e:
        return []


def save_gsheet_link(office, link):
    """بيحفظ لينك Google Sheets بتاع المكتب — بترجع (ok, message) عشان تبان المشكلة الحقيقية"""
    try:
        sheet = get_accounts_sheet()
        if not sheet:
            return False, "مش قادر أوصل لشيت الحسابات (accounts sheet) — تأكدي من صلاحيات الـ service account."

        headers = sheet.row_values(1)
        target = str(office).strip()

        # بندور صف صف على عمود "اسم المكتب" (العمود الأول) بدل get_all_records
        # عشان نتجنب مشاكل الهيدرز الفاضية/المكررة
        office_col_values = sheet.col_values(1)
        row_num = None
        for i, val in enumerate(office_col_values[1:], start=2):  # صف 1 = هيدر
            if str(val).strip() == target:
                row_num = i
                break

        if row_num is None:
            all_names = [str(v).strip() for v in office_col_values[1:]]
            return False, f"مش لاقي مكتب اسمه '{target}' في شيت الحسابات. الأسماء الموجودة فعلاً: {all_names}"

        # لو مفيش عمود للينك، ضيفيه (مع توسيع الشيت لو لازم)
        if "لينك الشيت" not in headers:
            col_num = len(headers) + 1
            if col_num > sheet.col_count:
                sheet.add_cols(col_num - sheet.col_count)
            sheet.update_cell(1, col_num, "لينك الشيت")
        else:
            col_num = headers.index("لينك الشيت") + 1

        sheet.update_cell(row_num, col_num, link)
        return True, "تم الحفظ بنجاح"

    except Exception as e:
        print(f"خطأ في حفظ اللينك: {e}")
        return False, f"خطأ: {e}"


def get_gsheet_link(office):
    """جيب لينك Google Sheets بتاع المكتب"""
    try:
        sheet = get_accounts_sheet()
        if not sheet:
            return None
        headers = sheet.row_values(1)
        if "لينك الشيت" not in headers:
            return None
        link_col = headers.index("لينك الشيت") + 1

        # بندور صف صف على عمود "اسم المكتب" (العمود الأول) بدل get_all_records
        # عشان نتجنب مشاكل الهيدرز الفاضية/المكررة اللي ممكن تطلع استثناء صامت
        office_col_values = sheet.col_values(1)
        target = str(office).strip()
        for i, val in enumerate(office_col_values[1:], start=2):  # صف 1 = هيدر
            if str(val).strip() == target:
                link = sheet.cell(i, link_col).value
                return link if link else None
        return None
    except Exception as e:
        print(f"خطأ في get_gsheet_link: {e}")
        return None


def extract_sheet_id(link):
    """بيستخرج الـ Sheet ID من اللينك"""
    import re
    match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', link)
    if match:
        return match.group(1)
    return None


def read_gsheet_as_excel(sheet_id):
    """بيقرا Google Sheet ويرجعه كـ BytesIO زي إكسيل"""
    try:
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(
            creds_dict,
            scopes=["https://spreadsheets.google.com/feeds",
                    "https://www.googleapis.com/auth/drive"]
        )
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(sheet_id)
        ws = spreadsheet.sheet1
        data = ws.get_all_values()

        if not data:
            return None, "الشيت فاضي!"

        wb = openpyxl.Workbook()
        wsheet = wb.active
        for row in data:
            wsheet.append(row)

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return out, None
    except Exception as e:
        return None, str(e)


def write_back_to_gsheet(sheet_id, wb):
    """بيكتب النتائج تاني في Google Sheet"""
    try:
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(
            creds_dict,
            scopes=["https://spreadsheets.google.com/feeds",
                    "https://www.googleapis.com/auth/drive"]
        )
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(sheet_id)
        ws = spreadsheet.sheet1
        wsheet = wb.active

        # حدث كل الخلايا
        data = []
        for row in wsheet.iter_rows(values_only=True):
            data.append([str(c) if c is not None else "" for c in row])

        ws.clear()
        ws.update(data)
        return True
    except Exception as e:
        return False


def get_results_sheet():
    """جيب أو أنشئ ورقة النتائج الموحدة (بديل رفع الملفات على Drive اللي مش شغال مع service account)"""
    try:
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(
            creds_dict,
            scopes=["https://spreadsheets.google.com/feeds",
                    "https://www.googleapis.com/auth/drive"]
        )
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(SHEET_ID)
        try:
            sheet = spreadsheet.worksheet("results")
        except:
            sheet = spreadsheet.add_worksheet("results", 2000, 4)
            sheet.append_row(["اسم المكتب", "اسم الطالب", "الحالة", "تاريخ التحديث"])
        return sheet
    except Exception as e:
        print(f"خطأ في ورقة النتائج: {e}")
        return None


def save_results_to_sheet(office, results):
    """بيحفظ آخر نتائج المكتب في ورقة 'results' — بيمسح القديم بتاع المكتب ده ويكتب الجديد"""
    try:
        sheet = get_results_sheet()
        if not sheet:
            return False, "مش قادر أوصل لورقة النتائج"

        all_values = sheet.get_all_values()
        headers = all_values[0] if all_values else ["اسم المكتب", "اسم الطالب", "الحالة", "تاريخ التحديث"]

        target = str(office).strip()
        # سيبي باقي المكاتب زي ما هي، وامسحي بس صفوف المكتب ده القديمة
        kept_rows = [row for row in all_values[1:] if row and str(row[0]).strip() != target]

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        new_rows = [[target, r.get("name", ""), r.get("status", ""), now] for r in results]

        final_data = [headers] + kept_rows + new_rows
        sheet.clear()
        sheet.update(final_data)
        return True, "تم حفظ النتائج"
    except Exception as e:
        print(f"خطأ في حفظ النتائج: {e}")
        return False, f"خطأ: {e}"


def search_results_in_sheet(office, name_query):
    """بيبحث عن طالب باسمه جوه آخر نتائج محفوظة للمكتب ده"""
    try:
        sheet = get_results_sheet()
        if not sheet:
            return [], "مش قادر أوصل لورقة النتائج"

        all_values = sheet.get_all_values()
        if len(all_values) <= 1:
            return [], "مفيش نتائج محفوظة لأي مكتب لسه — لازم تعملي '▶ ابدأ' مرة واحدة الأول."

        target_office = str(office).strip()
        query = str(name_query).strip()
        found = []
        office_has_any = False
        for row in all_values[1:]:
            if len(row) < 3:
                continue
            row_office, row_name, row_status = row[0], row[1], row[2]
            if str(row_office).strip() != target_office:
                continue
            office_has_any = True
            if query in str(row_name):
                found.append({"name": row_name, "status": row_status})

        if not office_has_any:
            return [], "مفيش نتائج محفوظة للمكتب ده لسه — لازم تعملي '▶ ابدأ' مرة واحدة الأول."
        return found, None
    except Exception as e:
        return [], f"خطأ: {e}"


def upload_to_drive(file_bytes, filename, office):
    """بيرفع الإكسيل على Google Drive — بترجع (ok, message) عشان تبان أي مشكلة فعلية"""
    try:
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseUpload
        import io as _io

        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(
            creds_dict,
            scopes=["https://www.googleapis.com/auth/drive"]
        )

        service = build("drive", "v3", credentials=creds)

        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        drive_filename = f"{office}_{now}_{filename}"

        file_metadata = {
            "name": drive_filename,
            "parents": [DRIVE_FOLDER_ID]
        }

        media = MediaIoBaseUpload(
            _io.BytesIO(file_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id",
            supportsAllDrives=True,
        ).execute()

        return True, "تم الرفع بنجاح"
    except Exception as e:
        print(f"خطأ في رفع الملف: {e}")
        return False, f"فشل الرفع على Drive: {e}"

# ==================== API ====================
BASE_URL = "https://apiadm.study-in-egypt.gov.eg/api"
SITE_URL = "https://admission.study-in-egypt.gov.eg"

HEADERS_BASE = {
    "accept": "application/json, text/plain, */*",
    "accept-language": "ar",
    "device": "CITIZEN",
    "origin": SITE_URL,
    "referer": SITE_URL + "/",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "content-type": "application/json",
}


def human_delay(min_sec=2, max_sec=5):
    """استنى وقت عشوائي زي إنسان"""
    time.sleep(random.uniform(min_sec, max_sec))


def api_login(email, password):
    """بيعمل لوجين ويرجع session جديدة فيها الكوكي"""
    try:
        session = requests.Session()
        session.headers.update(HEADERS_BASE)

        # استنى شوية قبل اللوجين
        human_delay(2, 4)

        res = session.post(
            f"{BASE_URL}/student/login",
            json={"email": email, "password": password},
            timeout=30
        )

        if res.status_code not in [200, 201]:
            return None, None, f"فشل اللوجين - كود: {res.status_code}"

        # استنى شوية بعد اللوجين
        human_delay(2, 3)

        csrf_token = res.json().get("token", "") or res.headers.get("x-csrf-token", "")

        return session, csrf_token, None

    except Exception as e:
        return None, None, str(e)


def api_logout(session):
    """تسجيل خروج"""
    try:
        session.post(
            f"{BASE_URL}/student/logout",
            json={"redirectUrl": SITE_URL},
            timeout=15
        )
        human_delay(1, 2)
    except:
        pass


def get_status(session, csrf_token):
    """بيجيب حالة أحدث طلب"""
    try:
        filter_obj = {
            "where": {},
            "limit": 10,
            "offset": 0,
            "order": "statusUpdatedAt DESC",
            "fields": ["serviceSlug", "ID", "createdAt", "statusUpdatedAt", "activityId", "activityName"]
        }

        headers = {}
        if csrf_token:
            headers["x-csrf-token"] = csrf_token

        # استنى شوية قبل الـ request
        human_delay(1, 3)

        res = session.get(
            f"{BASE_URL}/dynamic_services/inbox",
            params={"filter": json.dumps(filter_obj)},
            headers=headers,
            timeout=30
        )

        if res.status_code not in [200, 304]:
            return "", f"خطأ ({res.status_code})"

        data = res.json()
        results = data.get("result", [])

        if not results:
            return "", "مفيش طلبات"

        # أحدث طلب (مرتبين من الأحدث للأقدم)
        latest = results[0]

        # جدول ترجمة activityName للعربي الصح
        translations = {
            "قبول الفحص الفنى": "القبول المبدئي",
            "قبول الفحص الفني": "القبول المبدئي",
            "kb8ijfo8": "تم السداد",
            "تم السداد": "تم السداد",
            "تأكيد استلام الملف وصحة و اكتمال المستندات": "تأكيد استلام الملف وصحة واكتمال المستندات",
            "الانتظار مراجعة الطلب": "بانتظار مراجعة الطلب",
            "قبول من رئيس الادارة المركزية": "قبول من رئيس الإدارة المركزية",
        }
        raw_status = latest.get("activityName", "غير محدد")
        status = translations.get(raw_status, raw_status)
        app_id = str(latest.get("ID", ""))
        return app_id, status

    except Exception as e:
        return "", f"خطأ: {e}"


def find_excel_columns(ws):
    cols = {"name": None, "email": None, "password": None, "status": None}
    header_row_num = None
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), start=1):
        row_values = [str(c).strip() if c else "" for c in row]
        if any("يميل" in v or "mail" in v.lower() for v in row_values):
            header_row_num = row_idx
            for i, cell in enumerate(row_values):
                cell_lower = cell.lower()
                if any(k in cell for k in ["اسم", "الإسم", "الاسم"]) or "name" in cell_lower:
                    cols["name"] = i
                elif any(k in cell for k in ["يميل", "بريد"]) or "mail" in cell_lower:
                    cols["email"] = i
                elif any(k in cell for k in ["باسورد", "كلمة المرور", "password", "pass"]) or "pass" in cell_lower:
                    cols["password"] = i
                elif any(k in cell for k in ["حالة", "الحالة", "status"]):
                    cols["status"] = i
            break
    if header_row_num is None:
        raise Exception("مش لاقي هيدر الإكسيل!")
    if cols["email"] is None:
        raise Exception("مش لاقي عمود الإيميل!")
    if cols["password"] is None:
        raise Exception("مش لاقي عمود الباسورد!")
    return cols, header_row_num


# ==================== الواجهة ====================
st.set_page_config(
    menu_items={"Get help": None, "Report a bug": None, "About": None},
    page_title="ادرس في مصر - Agent",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ===== منع ترجمة المتصفح التلقائية (كانت بتلخبط عناصر الصفحة زي حقل كلمة المرور وزرار الرفع) =====
st.markdown("""
<meta name="google" content="notranslate">
""", unsafe_allow_html=True)

st.markdown("""
<style>
/* ===== Global ===== */
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"], .stApp {
    font-family: 'Cairo', sans-serif !important;
    direction: rtl;
    translate: no; /* منع ترجمة المتصفح التلقائية اللي بتلخبط الـ DOM */
}
.stApp { background: #f5f7fb; color: #111827; }
[data-testid="stAppViewContainer"] > .main { background: #f5f7fb; }
.block-container { max-width: 1180px; padding-top: 1.4rem; padding-bottom: 3rem; }
#MainMenu, footer, header, [data-testid="stToolbar"], [data-testid="stDecoration"],
[data-testid="stStatusWidget"], [data-testid="stSidebarNav"] { display:none !important; }

/* ===== Typography ===== */
h1, h2, h3, h4, p, label, span, div { font-family: 'Cairo', sans-serif !important; }
h1 { color:#111827 !important; font-size:32px !important; font-weight:800 !important; }
h2 { color:#111827 !important; font-size:23px !important; font-weight:800 !important; }
h3 { color:#111827 !important; font-size:19px !important; font-weight:700 !important; }
.stCaption, [data-testid="stCaptionContainer"] p { color:#6b7280 !important; }

/* ===== Top bar ===== */
.topbar {
    background:#ffffff;
    border:1px solid #e5e7eb;
    border-radius:16px;
    padding:13px 18px;
    display:flex;
    align-items:center;
    justify-content:space-between;
    box-shadow:0 2px 10px rgba(17,24,39,.04);
    margin-bottom:22px;
}
.brand { display:flex; align-items:center; gap:11px; }
.brand-icon {
    width:42px; height:42px; border-radius:12px; background:#eff6ff;
    display:flex; align-items:center; justify-content:center; font-size:22px;
}
.brand-title { font-size:18px; font-weight:800; color:#111827; }
.brand-sub { font-size:12px; color:#6b7280; margin-top:-2px; }

/* ===== Cards ===== */
.card {
    background:#fff; border:1px solid #e5e7eb; border-radius:18px;
    padding:22px; box-shadow:0 3px 14px rgba(17,24,39,.045);
    margin-bottom:18px;
}
.hero {
    background:linear-gradient(135deg,#ffffff 0%,#f8fbff 100%);
    border:1px solid #dbeafe; border-radius:20px; padding:25px 28px;
    box-shadow:0 4px 18px rgba(37,99,235,.06); margin-bottom:20px;
}
.hero-kicker { color:#6b7280; font-size:14px; font-weight:600; }
.hero-title { color:#111827; font-size:28px; font-weight:800; margin-top:2px; }
.hero-title strong { color:#2563eb; }
.hero-desc { color:#6b7280; font-size:14px; margin-top:3px; }

.stat-card { background:#fff; border:1px solid #e5e7eb; border-radius:16px; padding:17px 18px; }
.stat-label { color:#6b7280; font-size:12px; font-weight:600; }
.stat-value { color:#111827; font-size:21px; font-weight:800; margin-top:2px; }

.section-title { font-size:18px; font-weight:800; color:#111827; margin:5px 0 13px; }
.section-sub { color:#6b7280; font-size:13px; margin-top:-8px; margin-bottom:14px; }

/* ===== Inputs ===== */
.stTextInput label, .stFileUploader label, .stRadio > label, .stCheckbox label {
    color:#374151 !important; font-size:14px !important; font-weight:700 !important;
}
.stTextInput input {
    background:#fff !important; color:#111827 !important; border:1px solid #d1d5db !important;
    border-radius:10px !important; font-size:14px !important; min-height:44px;
}
.stTextInput input:focus { border-color:#2563eb !important; box-shadow:0 0 0 3px rgba(37,99,235,.10) !important; }
.stTextInput input::placeholder { color:#9ca3af !important; }

/* ===== Buttons ===== */
.stButton > button, .stDownloadButton > button {
    width:100%; min-height:44px; border-radius:10px !important; border:1px solid #2563eb !important;
    background:#2563eb !important; color:#fff !important; font-weight:700 !important;
    font-size:14px !important; box-shadow:0 3px 8px rgba(37,99,235,.16) !important;
    transition:.15s ease;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    background:#1d4ed8 !important; border-color:#1d4ed8 !important; transform:translateY(-1px);
}

/* Secondary buttons */
button[kind="secondary"] { background:#fff !important; color:#2563eb !important; }

/* ===== Tabs ===== */
button[data-baseweb="tab"] {
    color:#6b7280 !important; font-weight:700 !important; font-size:14px !important;
}
button[data-baseweb="tab"][aria-selected="true"] { color:#2563eb !important; }
[data-baseweb="tab-highlight"] { background:#2563eb !important; height:2px !important; }

/* ===== Radio / uploader ===== */
.stRadio div[role="radiogroup"] { gap:10px; }
.stRadio div[role="radiogroup"] label {
    background:#fff; border:1px solid #e5e7eb; border-radius:12px; padding:10px 14px;
}
[data-testid="stFileUploaderDropzone"] {
    background:#f8fafc !important; border:1.5px dashed #cbd5e1 !important; border-radius:14px !important;
}
[data-testid="stFileUploaderDropzone"] button {
    background:#fff !important; color:#2563eb !important; border:1px solid #bfdbfe !important;
    border-radius:8px !important;
}

/* ===== Alerts ===== */
div[data-testid="stAlert"] { border-radius:12px !important; border:1px solid #e5e7eb !important; }
div[data-testid="stAlert"] p { font-size:13px !important; font-weight:600 !important; }

/* ===== Search results ===== */
.result-card { background:#fff; border:1px solid #e5e7eb; border-radius:13px; padding:14px 16px; margin:8px 0; }
.result-name { color:#111827; font-size:15px; font-weight:800; }
.result-status { color:#2563eb; font-size:13px; font-weight:700; margin-top:2px; }

/* ===== Status badges ===== */
.status-badge { display:inline-block; padding:4px 10px; border-radius:999px; font-size:12px; font-weight:700; }
.status-ok { background:#ecfdf5; color:#15803d; }
.status-warn { background:#fffbeb; color:#b45309; }
.status-error { background:#fef2f2; color:#b91c1c; }
.status-info { background:#eff6ff; color:#1d4ed8; }

/* ===== Hide Streamlit input instructions / password visibility toggle =====
   استخدمنا selector مش معتمد على :has() عشان يشتغل على كل المتصفحات
   وعشان نتجنب ظهور زرار "Upload" مكرر بسبب ترجمة المتصفح أو أي DOM إضافي غريب */
[data-testid="InputInstructions"] { display:none !important; }
[data-testid="stTextInput"] button {
    display:none !important;
    visibility:hidden !important;
    width:0 !important;
    height:0 !important;
    padding:0 !important;
    margin:0 !important;
}

/* ===== Mobile ===== */
@media (max-width: 700px) {
    .block-container { padding: .8rem .7rem 2rem; }
    .hero-title { font-size:23px; }
    .topbar { padding:11px 13px; }
}
</style>
""", unsafe_allow_html=True)

# ===== Session state =====
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "is_admin" not in st.session_state:
    st.session_state.is_admin = False

# ===== Login =====
if not st.session_state.logged_in and not st.session_state.is_admin:
    st.markdown("""
    <div style="text-align:center; margin:42px 0 24px;">
        <div style="font-size:48px;">🎓</div>
        <div style="font-size:31px;font-weight:800;color:#111827;">ادرس في مصر</div>
        <div style="font-size:14px;color:#6b7280;margin-top:3px;">منصة متابعة طلبات الطلاب للمكاتب</div>
    </div>
    """, unsafe_allow_html=True)

    left, right = st.columns([1.15, 1], gap="large")
    with left:
        st.markdown("""
        <div style="padding:40px 20px 20px 10px;">
            <div style="font-size:14px;color:#2563eb;font-weight:800;margin-bottom:8px;">حل بسيط لإدارة المتابعة</div>
            <div style="font-size:31px;font-weight:800;color:#111827;line-height:1.35;">تابعي طلبات طلابك<br>من مكان واحد.</div>
            <div style="font-size:15px;color:#6b7280;line-height:1.9;margin-top:12px;max-width:480px;">
                حدّثي حالات الطلبات، اربطي Google Sheets، وابحثي عن أي طالب بسرعة بدون متابعة يدوية.
            </div>
            <div style="margin-top:22px;color:#374151;font-size:14px;line-height:2.2;">
                ✓ تحديث حالات الطلاب بشكل منظم<br>
                ✓ حفظ مصدر البيانات للمكتب<br>
                ✓ بحث سريع عن حالة أي طالب
            </div>
        </div>
        """, unsafe_allow_html=True)

    with right:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        tab1, tab2 = st.tabs(["تسجيل الدخول", "حساب جديد"])
        with tab1:
            st.markdown("<div style='font-size:22px;font-weight:800;color:#111827;margin:8px 0 3px;'>مرحبًا بعودتك 👋</div><div style='color:#6b7280;font-size:13px;margin-bottom:18px;'>سجلي دخولك لإدارة طلبات الطلاب</div>", unsafe_allow_html=True)
            username = st.text_input("اسم المكتب", key="login_user", placeholder="اكتبي اسم المكتب")
            password = st.text_input("كلمة المرور", type="password", key="login_pass", placeholder="اكتبي كلمة المرور")
            if st.button("تسجيل الدخول", key="login_btn"):
                if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
                    st.session_state.is_admin = True
                    st.rerun()
                else:
                    ok, msg = check_login(username, password)
                    if ok:
                        st.session_state.logged_in = True
                        st.session_state.office = username
                        log_to_sheet(username, "تسجيل دخول")
                        st.rerun()
                    else:
                        st.error(msg)
        with tab2:
            st.markdown("<div style='font-size:22px;font-weight:800;color:#111827;margin:8px 0 3px;'>إنشاء حساب</div><div style='color:#6b7280;font-size:13px;margin-bottom:18px;'>سجلي مكتبك لبدء استخدام الخدمة</div>", unsafe_allow_html=True)
            new_office = st.text_input("اسم المكتب", key="reg_office", placeholder="اسم المكتب")
            new_email = st.text_input("الإيميل", key="reg_email", placeholder="example@email.com")
            new_pass = st.text_input("كلمة المرور", type="password", key="reg_pass", placeholder="كلمة المرور")
            new_pass2 = st.text_input("تأكيد كلمة المرور", type="password", key="reg_pass2", placeholder="أعيدي كتابة كلمة المرور")
            if st.button("إنشاء الحساب", key="reg_btn"):
                if not new_office or not new_email or not new_pass:
                    st.error("املي كل الحقول!")
                elif new_pass != new_pass2:
                    st.error("كلمة المرور مش متطابقة!")
                else:
                    ok, msg = register_office(new_office, new_email, new_pass)
                    if ok:
                        st.success(msg)
                    else:
                        st.error(msg)
        st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

# ===== Top bar =====
st.markdown("""
<div class="topbar">
    <div class="brand">
        <div class="brand-icon">🎓</div>
        <div><div class="brand-title">ادرس في مصر</div><div class="brand-sub">لوحة متابعة الطلبات</div></div>
    </div>
    <div style="font-size:13px;color:#6b7280;">نظام متابعة المكاتب</div>
</div>
""", unsafe_allow_html=True)

# ===== Admin =====
if st.session_state.is_admin:
    st.markdown("<div class='hero'><div class='hero-kicker'>الإدارة</div><div class='hero-title'>لوحة الإدارة</div><div class='hero-desc'>إدارة الحسابات الجديدة ومتابعة حالة المكاتب.</div></div>", unsafe_allow_html=True)
    pending = get_pending_accounts()
    st.markdown(f"<div class='section-title'>الحسابات المعلقة <span style='color:#2563eb'>({len(pending)})</span></div>", unsafe_allow_html=True)
    if pending:
        for acc in pending:
            c1, c2, c3, c4 = st.columns([3, 3, 1, 1])
            c1.write(acc.get("اسم المكتب", ""))
            c2.write(acc.get("الإيميل", ""))
            if c3.button("قبول", key=f"approve_{acc.get('اسم المكتب')}"):
                approve_account(acc.get("اسم المكتب")); st.rerun()
            if c4.button("رفض", key=f"reject_{acc.get('اسم المكتب')}"):
                reject_account(acc.get("اسم المكتب")); st.rerun()
    else:
        st.info("لا توجد حسابات في انتظار الموافقة.")

    st.markdown("<div class='section-title' style='margin-top:28px;'>كل الحسابات</div>", unsafe_allow_html=True)
    try:
        sheet = get_accounts_sheet()
        if sheet:
            all_accounts = sheet.get_all_records()
            if all_accounts:
                df = pd.DataFrame(all_accounts).drop(columns=["الباسورد"], errors="ignore")
                st.dataframe(df, use_container_width=True, hide_index=True)
    except Exception:
        pass
    if st.button("تسجيل الخروج من الإدارة"):
        st.session_state.is_admin = False
        st.rerun()
    st.stop()

# ===== Main dashboard =====
_hour = datetime.now().hour
_greeting = "صباح الخير" if _hour < 12 else "مساء الخير"
office = st.session_state.office

st.markdown(f"""
<div class="hero">
    <div class="hero-kicker">{_greeting} 👋</div>
    <div class="hero-title">أهلاً بيك، <strong>{office}</strong></div>
    <div class="hero-desc">تابعي طلبات طلابك وحدّثي الحالات من مكان واحد.</div>
</div>
""", unsafe_allow_html=True)

saved_link = get_gsheet_link(office)

st.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)

# ===== Data source card =====
st.markdown("<div class='card'>", unsafe_allow_html=True)
st.markdown("<div class='section-title'>مصدر بيانات الطلاب</div><div class='section-sub'>اختاري الطريقة التي يحتوي بها ملف الطلاب.</div>", unsafe_allow_html=True)
source_options = ["📂 رفع ملف Excel", "🔗 ربط Google Sheets"]
source = st.radio("", source_options, horizontal=True, label_visibility="collapsed")

file_bytes = None
sheet_id_source = None
filename = ""

if source == "📂 رفع ملف Excel":
    uploaded = st.file_uploader("ارفع ملف Excel", type=["xlsx", "xls"], label_visibility="collapsed")
    if uploaded:
        file_bytes = uploaded.read()
        filename = uploaded.name
        st.success(f"تم اختيار الملف: {uploaded.name}")

else:
    if saved_link:
        st.markdown(f"<div style='background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;padding:11px 14px;color:#166534;font-size:13px;'>✓ Google Sheets متصل بالفعل لهذا المكتب</div>", unsafe_allow_html=True)
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        change = st.checkbox("تغيير رابط الشيت")
        if change:
            new_link = st.text_input("رابط Google Sheets", placeholder="الصق الرابط هنا")
            if st.button("حفظ الرابط"):
                sid_check = extract_sheet_id(new_link) if new_link else None
                if not new_link:
                    st.error("ادخلي الرابط أولاً!")
                elif not sid_check:
                    st.error("الرابط غير صحيح!")
                else:
                    ok, msg = save_gsheet_link(office, new_link)
                    if ok:
                        st.success("تم حفظ الرابط!"); st.rerun()
                    else:
                        st.error(msg)
        sheet_id_source = extract_sheet_id(saved_link)
    else:
        new_link = st.text_input("رابط Google Sheets", placeholder="https://docs.google.com/spreadsheets/d/...")
        if st.button("حفظ وربط الشيت"):
            sid = extract_sheet_id(new_link) if new_link else None
            if sid:
                ok, msg = save_gsheet_link(office, new_link)
                if ok:
                    st.success("تم حفظ الرابط بنجاح!"); st.rerun()
                else:
                    st.error(msg)
            else:
                st.error("الرابط غير صحيح!")

    if sheet_id_source:
        update_option = st.radio("طريقة التحديث", ["🔄 تحديث من الشيت الأونلاين", "📂 رفع Excel جديد بدلاً عنه"], horizontal=True)
        if update_option == "🔄 تحديث من الشيت الأونلاين":
            if st.button("تحميل بيانات الشيت"):
                with st.spinner("بيقرأ الشيت..."):
                    result, err = read_gsheet_as_excel(sheet_id_source)
                    if err:
                        st.error(f"خطأ: {err}"); st.stop()
                    file_bytes = result.read(); filename = "google_sheet"
                    st.session_state.pending_file_bytes = file_bytes
                    st.session_state.pending_filename = filename
                    st.success("تم جلب بيانات الشيت. البيانات جاهزة للتحديث.")
        else:
            uploaded2 = st.file_uploader("ارفع ملف Excel الجديد", type=["xlsx", "xls"], label_visibility="collapsed")
            if uploaded2:
                file_bytes = uploaded2.read(); filename = uploaded2.name
                st.session_state.pending_file_bytes = file_bytes
                st.session_state.pending_filename = filename
st.markdown('</div>', unsafe_allow_html=True)

if not file_bytes and st.session_state.get("pending_file_bytes"):
    file_bytes = st.session_state.pending_file_bytes
    filename = st.session_state.get("pending_filename", filename)

# ===== Processing =====
if file_bytes:
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.markdown("<div class='section-title'>تحديث حالات الطلاب</div><div class='section-sub'>اضغطي الزر لبدء فحص الطلبات وتحديث النتائج.</div>", unsafe_allow_html=True)
    if st.button("▶ تحديث حالات الطلاب", key="start_main"):
        log_to_sheet(office, "رفع ملف", filename)
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        ws = wb.active
        try:
            cols, header_row_num = find_excel_columns(ws)
        except Exception as e:
            st.error(str(e)); st.stop()

        rows_data = list(ws.iter_rows(min_row=header_row_num + 1, values_only=False))
        valid_rows = [r for r in rows_data if r[cols["email"]].value and r[cols["password"]].value]
        total = len(valid_rows)
        progress = st.progress(0)
        status_placeholder = st.empty()
        success = failed = 0

        for idx, row in enumerate(valid_rows):
            email = str(row[cols["email"]].value).strip()
            password = str(row[cols["password"]].value).strip()
            name = row[cols["name"]].value if cols["name"] is not None else ""
            status_placeholder.markdown(f"**جاري تحديث:** {name or email}<br><span style='color:#6b7280'>طالب {idx+1} من {total}</span>", unsafe_allow_html=True)
            session, csrf_token, err = api_login(email, password)
            if err or not session:
                if cols["status"] is not None: row[cols["status"]].value = "فشل تسجيل الدخول"
                failed += 1
            else:
                app_num, status = get_status(session, csrf_token)
                if cols["status"] is not None: row[cols["status"]].value = status
                api_logout(session)
                success += 1
            progress.progress((idx + 1) / max(total, 1))
            if idx < total - 1: human_delay(5, 10)

        out = io.BytesIO(); wb.save(out); out.seek(0)
        if sheet_id_source and source == "🔗 ربط Google Sheets":
            with st.spinner("بيحدث Google Sheets..."):
                write_back_to_gsheet(sheet_id_source, wb)
            st.success("تم تحديث Google Sheets تلقائيًا!")
        up_ok, up_msg = upload_to_drive(out.getvalue(), filename, office)
        if up_ok: st.caption("تم حفظ نسخة احتياطية على Drive")
        log_to_sheet(office, "اكتمل المعالجة", filename)
        st.session_state.pending_file_bytes = None
        st.session_state.pending_filename = ""
        st.session_state.last_results = [{"name": str(r[cols["name"]].value if cols["name"] is not None else ""), "status": str(r[cols["status"]].value if cols["status"] is not None else "")} for r in valid_rows]
        with st.spinner("بيحفظ النتائج للبحث..."):
            res_ok, res_msg = save_results_to_sheet(office, st.session_state.last_results)
        st.markdown(f"<div style='background:#ecfdf5;border:1px solid #bbf7d0;border-radius:14px;padding:16px;margin-top:15px;'><div style='font-size:17px;font-weight:800;color:#166534;'>اكتمل التحديث 🎉</div><div style='color:#166534;font-size:13px;margin-top:4px;'>إجمالي {total} طالب · نجح {success} · فشل {failed}</div></div>", unsafe_allow_html=True)
        st.download_button(label="⬇ تحميل ملف Excel المحدث", data=out, file_name="students_updated.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.markdown('</div>', unsafe_allow_html=True)

# ===== Search =====
st.markdown("<div class='card'>", unsafe_allow_html=True)
st.markdown("<div class='section-title'>البحث عن طالب</div><div class='section-sub'>اكتبي اسم الطالب لمعرفة آخر حالة محفوظة.</div>", unsafe_allow_html=True)
search_query = st.text_input("اسم الطالب", placeholder="مثال: Ahmed Mohamed", label_visibility="collapsed")
if search_query:
    with st.spinner("بيبحث..."):
        found, err = search_results_in_sheet(office, search_query)
    if err:
        st.warning(err)
    elif found:
        for r in found:
            status = r.get("status", "")
            st.markdown(f"<div class='result-card'><div class='result-name'>👤 {r.get('name','')}</div><div class='result-status'>● {status}</div></div>", unsafe_allow_html=True)
    else:
        st.info("مفيش طالب بالاسم ده.")
st.markdown('</div>', unsafe_allow_html=True)

# ===== Footer / logout =====
st.markdown("<div style='height:5px'></div>", unsafe_allow_html=True)
if st.button("تسجيل الخروج", key="logout_main"):
    log_to_sheet(st.session_state.get("office", ""), "تسجيل خروج")
    st.session_state.clear()
    st.rerun()
