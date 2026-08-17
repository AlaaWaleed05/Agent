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
    layout="centered"
)

st.markdown("""
<style>
    body { direction: rtl; }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    [data-testid="stToolbar"] {visibility: hidden;}
    [data-testid="stDecoration"] {visibility: hidden;}
    [data-testid="stStatusWidget"] {visibility: hidden;}
    .viewerBadge_container__r5tak {display: none;}
    .styles_viewerBadge__CvC9N {display: none;}
    #stDecoration {display: none;}
    .stApp > header {display: none;}
    [data-testid="stSidebarNav"] {display: none;}
    .css-1jc7ptx, .e1ewe7hr3, .viewerBadge_container__1QSob,
    .styles_viewerBadge__1yB5_, .viewerBadge_link__1S137,
    .viewerBadge_text__1JaDK, div[class*="viewerBadge"],
    div[class*="StatusWidget"], [data-testid="stBottom"] > div:last-child {display: none !important;}

    /* إخفاء رسالة "Press Enter to apply" اللي بتظهر تحت حقول الإدخال */
    [data-testid="InputInstructions"] {display: none !important;}
    div[data-baseweb="input"] + div small {display: none !important;}

    /* ==== خلفية متدرجة حيوية ومريحة للعين ==== */
    .stApp {
        background: linear-gradient(160deg, #1a1035 0%, #1e1450 35%, #142850 70%, #0f1e3d 100%);
    }

    [data-testid="stAppViewContainer"] > .main {
        background: transparent;
    }

    h1 {
        background: linear-gradient(90deg, #ff6b9d, #ffa36c, #7dd8ff);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        text-align: center;
        font-weight: 800 !important;
    }

    .stButton > button {
        background: linear-gradient(90deg, #ff6b9d, #7d5fff);
        color: #ffffff;
        font-weight: bold;
        border: none;
        border-radius: 10px;
        width: 100%;
        padding: 0.55em 0;
        transition: transform 0.15s ease, box-shadow 0.15s ease;
        box-shadow: 0 4px 14px rgba(125, 95, 255, 0.35);
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(255, 107, 157, 0.4);
    }

    /* ==== إضافات تحسين الوضوح: أحجام خط وألوان أوضح لكل العناصر ==== */

    /* نص عام في التطبيق */
    .stApp, .stApp p, .stApp span, .stApp label, .stApp div {
        color: #eef2fb;
        font-size: 17px;
    }

    /* العناوين الفرعية (subheader) */
    h2, h3, .stApp h2, .stApp h3 {
        color: #7dd8ff !important;
        font-size: 24px !important;
        font-weight: 700 !important;
        margin-top: 18px;
    }

    /* الديفايدر أوضح شوية */
    hr { border-color: rgba(255,255,255,0.15) !important; }

    /* ليبلز حقول الإدخال (اسم المكتب، الباسورد، إلخ) */
    .stTextInput label, .stTextInput > label p,
    .stRadio label, .stRadio > label p,
    .stCheckbox label, .stCheckbox > label p,
    .stFileUploader label, .stFileUploader > label p {
        color: #ffffff !important;
        font-size: 18px !important;
        font-weight: 600 !important;
    }

    /* حقول الكتابة نفسها */
    .stTextInput input {
        background-color: rgba(255,255,255,0.92) !important;
        color: #0f1923 !important;
        font-size: 17px !important;
        border: 1px solid rgba(125, 216, 255, 0.35) !important;
        border-radius: 10px !important;
    }
    .stTextInput input:focus {
        border: 1px solid #7dd8ff !important;
        box-shadow: 0 0 0 3px rgba(125, 216, 255, 0.2) !important;
    }
    .stTextInput input::placeholder { color: #5c6b85 !important; }

    /* زرار "Browse files" وكل نص أوضة رفع الملفات */
    [data-testid="stFileUploaderDropzone"] button {
        color: #0f1923 !important;
        background-color: rgba(255,255,255,0.92) !important;
        font-weight: 700 !important;
        border-radius: 8px !important;
    }
    [data-testid="stFileUploaderDropzone"] button span,
    [data-testid="stFileUploaderDropzone"] button p {
        color: #0f1923 !important;
    }

    /* خيارات الراديو (نص الاختيارات) */
    .stRadio div[role="radiogroup"] label p { color: #eef2fb !important; font-size: 16px !important; }

    /* التابات (تسجيل الدخول / حساب جديد) */
    button[data-baseweb="tab"] {
        font-size: 17px !important;
        font-weight: 700 !important;
        color: #9fb3d9 !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        color: #ff6b9d !important;
    }
    [data-baseweb="tab-highlight"] {
        background: linear-gradient(90deg, #ff6b9d, #7d5fff) !important;
    }

    /* رسايل النجاح/التحذير/الخطأ/المعلومات — تباين أوضح */
    div[data-testid="stAlert"] {
        border-radius: 12px !important;
        border: 1px solid rgba(255,255,255,0.12) !important;
    }
    div[data-testid="stAlert"] p { font-size: 16px !important; font-weight: 600 !important; }
    div[data-testid="stAlertContentSuccess"] p,
    div[data-testid="stAlertContentInfo"] p { color: #ffffff !important; }

    /* التعليقات الصغيرة (caption) */
    .stApp small, [data-testid="stCaptionContainer"] p {
        color: #9fb3d9 !important;
        font-size: 14px !important;
    }

    /* الجداول (dataframe) */
    .stDataFrame, .stDataFrame * { color: #0f1923 !important; }

    /* الكود/لوج المعالجة */
    .stApp code, pre code { font-size: 15px !important; }

    /* بطاقة الترحيب */
    .welcome-hero {
        background: linear-gradient(120deg, rgba(255,107,157,0.18), rgba(125,95,255,0.18), rgba(125,216,255,0.15));
        border: 1px solid rgba(255,255,255,0.12);
        border-radius: 18px;
        padding: 22px 26px;
        margin-bottom: 10px;
        text-align: center;
    }
    .welcome-hero .greet {
        font-size: 15px;
        color: #9fb3d9;
        margin-bottom: 4px;
    }
    .welcome-hero .office-name {
        font-size: 28px;
        font-weight: 800;
        background: linear-gradient(90deg, #ff6b9d, #7d5fff, #7dd8ff);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }
    .welcome-hero .tagline {
        font-size: 14px;
        color: #c9d4ec;
        margin-top: 6px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("<h1>🎓 ادرس في مصر</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center;color:#9fb3d9'>Agent تتبع حالة الطلبات</p>",
            unsafe_allow_html=True)
st.divider()

# ==================== لوجين ====================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "is_admin" not in st.session_state:
    st.session_state.is_admin = False

if not st.session_state.logged_in and not st.session_state.is_admin:

    tab1, tab2 = st.tabs(["🔑 تسجيل الدخول", "📝 حساب جديد"])

    with tab1:
        username = st.text_input("اسم المكتب", key="login_user")
        password = st.text_input("كلمة المرور", type="password", key="login_pass")
        if st.button("دخول", key="login_btn"):
            # تحقق لو Admin
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
        st.markdown("##### سجلي مكتبك عشان تبدأ تستخدم الخدمة")
        new_office = st.text_input("اسم المكتب", key="reg_office")
        new_email  = st.text_input("الإيميل", key="reg_email")
        new_pass   = st.text_input("كلمة المرور", type="password", key="reg_pass")
        new_pass2  = st.text_input("تأكيد كلمة المرور", type="password", key="reg_pass2")
        if st.button("تسجيل", key="reg_btn"):
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

    st.stop()

# ==================== صفحة الأدمن ====================
if st.session_state.is_admin:
    st.subheader("🛡️ لوحة الإدارة")

    pending = get_pending_accounts()
    if pending:
        st.markdown(f"**{len(pending)} حساب في انتظار الموافقة:**")
        for acc in pending:
            col1, col2, col3, col4 = st.columns([3, 2, 1, 1])
            col1.write(acc.get("اسم المكتب",""))
            col2.write(acc.get("الإيميل",""))
            if col3.button("✅ قبول", key=f"approve_{acc.get('اسم المكتب')}"):
                approve_account(acc.get("اسم المكتب"))
                st.rerun()
            if col4.button("❌ رفض", key=f"reject_{acc.get('اسم المكتب')}"):
                reject_account(acc.get("اسم المكتب"))
                st.rerun()
    else:
        st.success("مفيش حسابات في انتظار الموافقة!")

    st.divider()

    # كل الحسابات
    st.markdown("**كل الحسابات:**")
    try:
        sheet = get_accounts_sheet()
        if sheet:
            all_accounts = sheet.get_all_records()
            if all_accounts:
                import pandas as pd
                df = pd.DataFrame(all_accounts)
                df = df.drop(columns=["الباسورد"], errors="ignore")
                st.dataframe(df, use_container_width=True)
    except:
        pass

    if st.button("خروج من الإدارة"):
        st.session_state.is_admin = False
        st.rerun()
    st.stop()

# ==================== الصفحة الرئيسية ====================
_hour = datetime.now().hour
if _hour < 12:
    _greeting = "☀️ صباح الخير"
elif _hour < 18:
    _greeting = "🌤️ مساء الخير"
else:
    _greeting = "🌙 مساء الخير"

st.markdown(f"""
<div class="welcome-hero">
    <div class="greet">{_greeting}، أهلاً بيك تاني</div>
    <div class="office-name">🏢 {st.session_state.office}</div>
    <div class="tagline">جاهزين نتابع طلبات طلابك خطوة بخطوة ✨</div>
</div>
""", unsafe_allow_html=True)

# ==================== تتبع الحالات ====================
st.subheader("🔍 تتبع حالات الطلاب")

# ---- خيار المصدر ----
saved_link = get_gsheet_link(st.session_state.office)

source_options = ["📂 رفع ملف إكسيل", "🔗 ربط Google Sheets"]
source = st.radio("اختاري مصدر البيانات:", source_options, horizontal=True)

file_bytes = None
sheet_id_source = None
filename = ""

if source == "📂 رفع ملف إكسيل":
    uploaded = st.file_uploader("ارفع ملف الإكسيل", type=["xlsx", "xls"])
    if uploaded:
        file_bytes = uploaded.read()
        filename = uploaded.name

elif source == "🔗 ربط Google Sheets":
    if saved_link:
        st.success(f"✅ متربط بشيت: `{saved_link[:60]}...`")
        change = st.checkbox("تغيير اللينك")
        if change:
            new_link = st.text_input("اللينك الجديد:")
            if st.button("حفظ اللينك"):
                sid_check = extract_sheet_id(new_link) if new_link else None
                if not new_link:
                    st.error("ادخلي اللينك الأول!")
                elif not sid_check:
                    st.error("اللينك مش صح!")
                else:
                    ok, msg = save_gsheet_link(st.session_state.office, new_link)
                    if ok:
                        st.success("✅ تم حفظ اللينك!")
                        st.rerun()
                    else:
                        st.error(f"❌ {msg}")
        sheet_id_source = extract_sheet_id(saved_link)
    else:
        new_link = st.text_input("الصق لينك Google Sheets هنا:")
        if st.button("💾 حفظ اللينك"):
            if new_link:
                sid = extract_sheet_id(new_link)
                if sid:
                    ok, msg = save_gsheet_link(st.session_state.office, new_link)
                    if ok:
                        st.success("✅ تم حفظ اللينك!")
                        st.rerun()
                    else:
                        st.error(f"❌ {msg}")
                else:
                    st.error("اللينك مش صح!")
            else:
                st.error("ادخلي اللينك الأول!")

    if sheet_id_source:
        update_option = st.radio(
            "عايزة إيه؟",
            ["🔄 تحديث من الشيت الأونلاين", "📂 رفع إكسيل جديد بدلاً عنه"],
            horizontal=True
        )
        if update_option == "🔄 تحديث من الشيت الأونلاين":
            if st.button("▶ ابدأ التحديث"):
                with st.spinner("بيقرا الشيت..."):
                    result, err = read_gsheet_as_excel(sheet_id_source)
                    if err:
                        st.error(f"خطأ: {err}")
                        st.stop()
                    file_bytes = result.read()
                    filename = "google_sheet"
                    # احفظيها في session_state عشان متضيعش لما الصفحة تعمل rerun
                    st.session_state.pending_file_bytes = file_bytes
                    st.session_state.pending_filename = filename
                    st.success("✅ اتجابت بيانات الشيت! دوسي '▶ ابدأ' تحت عشان تبدأي المعالجة.")
        else:
            uploaded2 = st.file_uploader("ارفع الإكسيل الجديد", type=["xlsx", "xls"])
            if uploaded2:
                file_bytes = uploaded2.read()
                filename = uploaded2.name
                st.session_state.pending_file_bytes = file_bytes
                st.session_state.pending_filename = filename

# لو مفيش file_bytes في المتغير المحلي (زي بعد rerun بتاع زرار "ابدأ التحديث")، رجعيه من session_state
if not file_bytes and st.session_state.get("pending_file_bytes"):
    file_bytes = st.session_state.pending_file_bytes
    filename = st.session_state.get("pending_filename", filename)

if file_bytes and st.button("▶ ابدأ"):
    log_to_sheet(st.session_state.office, "رفع ملف", filename)
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws = wb.active

    try:
        cols, header_row_num = find_excel_columns(ws)
    except Exception as e:
        st.error(str(e))
        st.stop()

    log_area = st.empty()
    progress = st.progress(0)

    rows_data  = list(ws.iter_rows(min_row=header_row_num + 1, values_only=False))
    valid_rows = [r for r in rows_data
                  if r[cols["email"]].value and r[cols["password"]].value]
    total = len(valid_rows)

    log_lines = []
    success = failed = 0

    for idx, row in enumerate(valid_rows):
        email    = str(row[cols["email"]].value).strip()
        password = str(row[cols["password"]].value).strip()
        name     = row[cols["name"]].value if cols["name"] is not None else ""

        log_lines.append(f"👤 {name} | {email}")
        log_area.code("\n".join(log_lines))

        # لوجين
        session, csrf_token, err = api_login(email, password)

        if err or not session:
            log_lines.append(f"   ❌ فشل اللوجين: {err}\n")
            if cols["status"] is not None:
                row[cols["status"]].value = "فشل تسجيل الدخول"
            failed += 1
        else:
            # جيب الحالة
            app_num, status = get_status(session, csrf_token)
            if cols["status"] is not None:
                row[cols["status"]].value = status
            log_lines.append(f"   ✅ {status}")

            # تسجيل خروج
            api_logout(session)
            log_lines.append(f"   🚪 تم تسجيل الخروج\n")
            success += 1

        log_area.code("\n".join(log_lines))
        progress.progress((idx + 1) / total)

        # delay بين كل طالب والتاني (5-10 ثواني)
        if idx < total - 1:
            human_delay(5, 10)

    # احفظ الإكسيل
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

    log_lines.append("─" * 40)
    log_lines.append(f"✅ خلصنا! إجمالي: {total} | نجح: {success} | فشل: {failed}")
    log_area.code("\n".join(log_lines))

    # لو مصدره Google Sheet، اكتب النتائج فيه أوتوماتيك
    if sheet_id_source and source == "🔗 ربط Google Sheets":
        with st.spinner("بيحدث الشيت الأونلاين..."):
            write_back_to_gsheet(sheet_id_source, wb)
        st.info("✅ تم تحديث Google Sheets أوتوماتيك!")

    # ارفع الإكسيل على Drive (اختياري/احتياطي — ممكن يفشل مع service account، مش هيوقف التطبيق)
    up_ok, up_msg = upload_to_drive(out.getvalue(), filename, st.session_state.office)
    if up_ok:
        st.caption("✅ اتحفظت نسخة احتياطية على Drive")
    log_to_sheet(st.session_state.office, "اكتمل المعالجة", filename)

    # امسحي الملف المؤقت من session_state عشان مايتشغلش تاني بالغلط
    st.session_state.pending_file_bytes = None
    st.session_state.pending_filename = ""

    # احفظ النتائج للبحث
    st.session_state.last_results = [
        {
            "name": str(r[cols["name"]].value if cols["name"] is not None else ""),
            "status": str(r[cols["status"]].value if cols["status"] is not None else ""),
        }
        for r in valid_rows
    ]

    # احفظ النتائج في شيت "results" عشان خاصية البحث تلاقيها لاحقًا
    with st.spinner("بيحفظ النتائج عشان البحث..."):
        res_ok, res_msg = save_results_to_sheet(st.session_state.office, st.session_state.last_results)
    if res_ok:
        st.caption("✅ النتائج جاهزة للبحث")
    else:
        st.warning(f"⚠️ اتعملت المعالجة بنجاح بس حفظ النتائج للبحث فشل: {res_msg}")

    st.success("خلصنا! حملي الإكسيل المحدث 👇")
    st.download_button(
        label="⬇ تحميل الإكسيل المحدث",
        data=out,
        file_name="students_updated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ==================== بحث ====================
st.divider()
st.subheader("🔍 بحث باسم الطالب")

search_query = st.text_input("اكتبي اسم الطالب:")
if search_query:
    with st.spinner("بيبحث..."):
        found, err = search_results_in_sheet(st.session_state.office, search_query)
        if err:
            st.warning(f"⚠️ {err}")
        elif found:
            for r in found:
                st.success(f"👤 **{r.get('name','')}** — {r.get('status','')}")
        else:
            st.warning("مفيش طالب بالاسم ده!")

st.divider()

if st.button("خروج"):
    log_to_sheet(st.session_state.get("office",""), "تسجيل خروج")
    st.session_state.clear()
    st.rerun()
