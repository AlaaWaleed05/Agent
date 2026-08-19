"""
auto_update.py
سكريبت مستقل (بره Streamlit) بيشتغل من GitHub Actions على جدول زمني ثابت.
بيدور على المكاتب اللي فعّلت "تحديث تلقائي"، ولو حان وقت تحديثها (حسب عدد
الساعات اللي اختاروها)، بيجيب آخر ملف محفوظ ليهم (Google Sheet مربوط أو
آخر نسخة على Drive) وبيحدّث حالات الطلاب فيه، ثم يحفظ النتائج، وبيبعت
إيميل للمكتب (على نفس إيميل التسجيل) بس لو في طالب اتغيرت حالته فعليًا.

الفرق الأساسي عن app.py: الملف ده مالوش أي تبعية على Streamlit (لأنه بيشتغل
من غير أي حد فاتح الموقع)، وبياخد بيانات الاعتماد (service account + إيميل
الإرسال) من متغيرات بيئة بدل st.secrets.
"""

import os
import io
import json
import time
import random
import requests
import openpyxl
import gspread
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials

SHEET_ID = "1BlFdtY-7ZIF1y2GwVosxlG9r7nK5xqYeW6yiIjPI_9U"
DRIVE_FOLDER_ID = "12L_qSHBnW4-tfQZRteynInWNBAML016f"

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


# ==================== Credentials ====================

def get_creds():
    creds_json = os.environ["GCP_SERVICE_ACCOUNT_JSON"]
    creds_dict = json.loads(creds_json)
    return Credentials.from_service_account_info(
        creds_dict,
        scopes=[
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
        ],
    )


def get_gspread_client():
    return gspread.authorize(get_creds())


# ==================== API helpers (نفس منطق app.py) ====================

def human_delay(min_sec=2, max_sec=5):
    time.sleep(random.uniform(min_sec, max_sec))


def api_login(email, password):
    try:
        session = requests.Session()
        session.headers.update(HEADERS_BASE)
        human_delay(2, 4)
        res = session.post(
            f"{BASE_URL}/student/login",
            json={"email": email, "password": password},
            timeout=30,
        )
        if res.status_code not in [200, 201]:
            return None, None, f"فشل اللوجين - كود: {res.status_code}"
        human_delay(2, 3)
        csrf_token = res.json().get("token", "") or res.headers.get("x-csrf-token", "")
        return session, csrf_token, None
    except Exception as e:
        return None, None, str(e)


def api_logout(session):
    try:
        session.post(f"{BASE_URL}/student/logout", json={"redirectUrl": SITE_URL}, timeout=15)
        human_delay(1, 2)
    except Exception:
        pass


def get_status(session, csrf_token):
    try:
        filter_obj = {
            "where": {},
            "limit": 10,
            "offset": 0,
            "order": "statusUpdatedAt DESC",
            "fields": ["serviceSlug", "ID", "createdAt", "statusUpdatedAt", "activityId", "activityName"],
        }
        headers = {}
        if csrf_token:
            headers["x-csrf-token"] = csrf_token
        human_delay(1, 3)
        res = session.get(
            f"{BASE_URL}/dynamic_services/inbox",
            params={"filter": json.dumps(filter_obj)},
            headers=headers,
            timeout=30,
        )
        if res.status_code not in [200, 304]:
            return "", f"خطأ ({res.status_code})"
        data = res.json()
        results = data.get("result", [])
        if not results:
            return "", "مفيش طلبات"
        latest = results[0]
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
        return str(latest.get("ID", "")), status
    except Exception as e:
        return "", f"خطأ: {e}"


def find_excel_columns(ws):
    cols = {"name": None, "email": None, "password": None, "status": None}
    header_row_num = None
    header_len = 0
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), start=1):
        row_values = [str(c).strip() if c else "" for c in row]
        if any("يميل" in v or "mail" in v.lower() for v in row_values):
            header_row_num = row_idx
            header_len = len(row_values)

            # أولوية: عمود اسمه "حالة الطلب الجديدة" تحديدًا - ده العمود اللي هنكتب فيه التحديث،
            # عشان ميتلخبطش مع أي عمود "حالة الطلب" أصلي موجود أصلاً في ملف المكتب
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

    # لو عمود "حالة الطلب الجديدة" مش موجود أصلاً، نضيفه تلقائيًا في آخر الشيت
    if cols["status"] is None:
        new_col_index = header_len
        ws.cell(row=header_row_num, column=new_col_index + 1, value="حالة الطلب الجديدة")
        cols["status"] = new_col_index

    return cols, header_row_num


def extract_sheet_id(link):
    import re
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", link)
    return match.group(1) if match else None


def extract_gid(link):
    """بيستخرج رقم الـ gid (بيحدد التبويب بالظبط جوه الملف) من اللينك لو موجود"""
    import re
    match = re.search(r"[?#&]gid=(\d+)", link)
    return int(match.group(1)) if match else None


def get_target_worksheet(spreadsheet, link):
    """بيرجع التبويب المطابق بالظبط للينك المحفوظ (حسب gid)، أو أول تبويب لو مفيش gid"""
    gid = extract_gid(link)
    if gid is not None:
        try:
            for ws in spreadsheet.worksheets():
                if ws.id == gid:
                    return ws
        except Exception:
            pass
    return spreadsheet.sheet1


# ==================== Data source (Google Sheet المربوط أو آخر نسخة Drive) ====================

def get_office_source_workbook(client, office, gsheet_link):
    """بيرجع (workbook, source_type, source_link) — source_type: 'gsheet' أو 'drive' أو None.
    لو gsheet، بيرجع اللينك الكامل (مش الـ ID بس) عشان write_back_to_gsheet يقدر يحدد نفس التبويب بالظبط."""
    if gsheet_link:
        sid = extract_sheet_id(gsheet_link)
        if sid:
            try:
                spreadsheet = client.open_by_key(sid)
                ws = get_target_worksheet(spreadsheet, gsheet_link)
                data = ws.get_all_values()
                if data:
                    wb = openpyxl.Workbook()
                    wsheet = wb.active
                    for row in data:
                        wsheet.append(row)
                    return wb, "gsheet", gsheet_link
            except Exception as e:
                print(f"تعذر قراءة الشيت المربوط للمكتب {office}: {e}")

    # لو مفيش شيت مربوط أو فشلت قراءته، نرجع لآخر نسخة محفوظة على Drive
    try:
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseDownload

        creds = get_creds()
        service = build("drive", "v3", credentials=creds)
        safe_office = str(office).replace("\\", "\\\\").replace("'", "\\'")
        results = (
            service.files()
            .list(
                q=f"'{DRIVE_FOLDER_ID}' in parents and name contains '{safe_office}' and trashed=false",
                orderBy="createdTime desc",
                pageSize=1,
                fields="files(id, name, createdTime)",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            )
            .execute()
        )
        files = results.get("files", [])
        if not files:
            return None, None, None

        file_id = files[0]["id"]
        request = service.files().get_media(fileId=file_id)
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        buf.seek(0)
        wb = openpyxl.load_workbook(buf)
        return wb, "drive", None
    except Exception as e:
        print(f"تعذر جلب آخر نسخة من Drive للمكتب {office}: {e}")
        return None, None, None


def upload_backup_to_drive(file_bytes, filename, office):
    try:
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseUpload

        creds = get_creds()
        service = build("drive", "v3", credentials=creds)
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        drive_filename = f"{office}_{now}_{filename}"
        file_metadata = {"name": drive_filename, "parents": [DRIVE_FOLDER_ID]}
        media = MediaIoBaseUpload(
            io.BytesIO(file_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        service.files().create(body=file_metadata, media_body=media, fields="id", supportsAllDrives=True).execute()
    except Exception as e:
        print(f"فشل رفع نسخة احتياطية للمكتب {office}: {e}")


def write_back_to_gsheet(client, link, wb):
    """بيكتب النتائج تاني في نفس التبويب المحدد في اللينك (حسب gid)"""
    try:
        sheet_id = extract_sheet_id(link)
        if not sheet_id:
            print("فشل تحديث Google Sheet: الرابط غير صحيح")
            return
        spreadsheet = client.open_by_key(sheet_id)
        ws = get_target_worksheet(spreadsheet, link)
        wsheet = wb.active
        data = [[str(c) if c is not None else "" for c in row] for row in wsheet.iter_rows(values_only=True)]
        ws.clear()
        ws.update(data)
    except Exception as e:
        print(f"فشل تحديث Google Sheet: {e}")


def get_previous_results(client, office):
    """بيرجع dict {اسم الطالب: الحالة} بناءً على آخر نتائج محفوظة للمكتب قبل التحديث الحالي"""
    try:
        spreadsheet = client.open_by_key(SHEET_ID)
        try:
            sheet = spreadsheet.worksheet("results")
        except Exception:
            return {}
        all_values = sheet.get_all_values()
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
        print(f"تعذر جلب النتائج السابقة للمكتب {office}: {e}")
        return {}


def compute_changes(previous_results, new_results):
    """بيقارن الحالات القديمة بالجديدة ويرجع بس الطلاب اللي حالتهم اتغيرت (أو طلاب جداد)"""
    changes = []
    for r in new_results:
        name = str(r.get("name", "")).strip()
        new_status = str(r.get("status", "")).strip()
        if not name:
            continue
        old_status = previous_results.get(name)
        if old_status is None:
            changes.append({"name": name, "new_status": new_status, "old_status": None})
        elif old_status != new_status:
            changes.append({"name": name, "new_status": new_status, "old_status": old_status})
    return changes


def build_email_message(office, changes):
    """بيبني نص الإيميل: قايمة مرقّمة بالاسم والحالة الجديدة بس (من غير الحالة القديمة)"""
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


def send_email_notification(to_email, subject, body):
    """بيبعت إيميل عن طريق Gmail SMTP - مجاني بالكامل، محتاج بس إيميل Gmail وApp Password"""
    import smtplib
    from email.mime.text import MIMEText

    sender_email = os.environ.get("SENDER_EMAIL")
    sender_password = os.environ.get("SENDER_APP_PASSWORD")

    if not sender_email or not sender_password:
        print("إعدادات الإيميل مش مكتملة على GitHub Secrets - تم تخطي الإرسال.")
        return
    if not to_email:
        print("مفيش إيميل محفوظ للمكتب - تم تخطي الإرسال.")
        return

    try:
        msg = MIMEText(body, "plain", "utf-8")
        msg["Subject"] = subject
        msg["From"] = sender_email
        msg["To"] = to_email

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, [to_email], msg.as_string())
        print(f"تم إرسال إشعار بالإيميل إلى {to_email}")
    except Exception as e:
        print(f"خطأ في إرسال الإيميل: {e}")


def save_results_to_sheet(client, office, results):
    try:
        spreadsheet = client.open_by_key(SHEET_ID)
        try:
            sheet = spreadsheet.worksheet("results")
        except Exception:
            sheet = spreadsheet.add_worksheet("results", 2000, 4)
            sheet.append_row(["اسم المكتب", "اسم الطالب", "الحالة", "تاريخ التحديث"])

        all_values = sheet.get_all_values()
        headers = all_values[0] if all_values else ["اسم المكتب", "اسم الطالب", "الحالة", "تاريخ التحديث"]
        target = str(office).strip()
        kept_rows = [row for row in all_values[1:] if row and str(row[0]).strip() != target]
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        new_rows = [[target, r.get("name", ""), r.get("status", ""), now] for r in results]
        final_data = [headers] + kept_rows + new_rows
        sheet.clear()
        sheet.update(final_data)
    except Exception as e:
        print(f"فشل حفظ نتائج المكتب {office}: {e}")


# ==================== Core logic ====================

def process_office(client, office, gsheet_link, office_email):
    print(f"--- بدء تحديث تلقائي للمكتب: {office} ---")
    wb, source_type, source_link = get_office_source_workbook(client, office, gsheet_link)
    if wb is None:
        print(f"لا يوجد ملف محفوظ للمكتب {office}، تم التخطي.")
        return

    ws = wb.active
    try:
        cols, header_row_num = find_excel_columns(ws)
    except Exception as e:
        print(f"خطأ في قراءة أعمدة ملف المكتب {office}: {e}")
        return

    # هات النتائج القديمة قبل ما نبدأ نكتب فوقها، عشان نقدر نقارن بعدين
    previous_results = get_previous_results(client, office)

    rows_data = list(ws.iter_rows(min_row=header_row_num + 1, values_only=False))
    valid_rows = [r for r in rows_data if r[cols["email"]].value and r[cols["password"]].value]

    for row in valid_rows:
        email = str(row[cols["email"]].value).strip()
        password = str(row[cols["password"]].value).strip()
        session, csrf_token, err = api_login(email, password)
        if err or not session:
            if cols["status"] is not None:
                row[cols["status"]].value = "فشل تسجيل الدخول"
        else:
            _, status = get_status(session, csrf_token)
            if cols["status"] is not None:
                row[cols["status"]].value = status
            api_logout(session)
        human_delay(5, 10)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

    if source_type == "gsheet" and source_link:
        write_back_to_gsheet(client, source_link, wb)

    upload_backup_to_drive(out.getvalue(), "auto_update.xlsx", office)

    results = [
        {
            "name": str(r[cols["name"]].value if cols["name"] is not None else ""),
            "status": str(r[cols["status"]].value if cols["status"] is not None else ""),
        }
        for r in valid_rows
    ]
    save_results_to_sheet(client, office, results)

    # لو أول مرة نجمع نتائج للمكتب ده، منبعتش إشعار (كل الحالات هتبان "جديدة" وده هيغرق المكتب بإيميل ضخم أول تشغيل)
    if previous_results:
        changes = compute_changes(previous_results, results)
        if changes:
            if office_email:
                subject = f"تحديث حالات الطلاب - {office}"
                body = build_email_message(office, changes)
                send_email_notification(office_email, subject, body)
            else:
                print(f"في {len(changes)} تغيير للمكتب {office}، بس مفيش إيميل محفوظ - تم تخطي الإرسال.")
        else:
            print(f"مفيش أي تغيير في حالات المكتب {office} - مفيش إشعار.")
    else:
        print(f"أول تحديث تلقائي للمكتب {office} - تم حفظ النتائج بدون إرسال إشعار.")

    print(f"--- اكتمل تحديث المكتب: {office} ({len(valid_rows)} طالب) ---")


def main():
    client = get_gspread_client()
    spreadsheet = client.open_by_key(SHEET_ID)
    accounts_sheet = spreadsheet.worksheet("accounts")
    records = accounts_sheet.get_all_records()
    headers = accounts_sheet.row_values(1)

    now = datetime.now()

    for i, r in enumerate(records, start=2):  # صف 1 = هيدر
        office = r.get("اسم المكتب", "")
        if not office:
            continue
        if r.get("الحالة", "") != "approved":
            continue
        if r.get("تحديث تلقائي", "لا") != "نعم":
            continue

        try:
            interval = int(r.get("كل كام ساعة", 12) or 12)
        except Exception:
            interval = 12

        last_update_str = r.get("آخر تحديث تلقائي", "")
        due = True
        if last_update_str:
            try:
                last_update = datetime.strptime(last_update_str, "%Y-%m-%d %H:%M:%S")
                due = now - last_update >= timedelta(hours=interval)
            except Exception:
                due = True

        if not due:
            print(f"المكتب {office}: لسه معملوش موعد التحديث (كل {interval} ساعة).")
            continue

        gsheet_link = r.get("لينك الشيت", "")
        office_email = str(r.get("الإيميل", "") or "").strip()
        try:
            process_office(client, office, gsheet_link, office_email)
            if "آخر تحديث تلقائي" not in headers:
                col_num = len(headers) + 1
                accounts_sheet.update_cell(1, col_num, "آخر تحديث تلقائي")
                headers = accounts_sheet.row_values(1)
            col_num = headers.index("آخر تحديث تلقائي") + 1
            accounts_sheet.update_cell(i, col_num, now.strftime("%Y-%m-%d %H:%M:%S"))
        except Exception as e:
            print(f"خطأ أثناء تحديث المكتب {office}: {e}")


if __name__ == "__main__":
    main()
