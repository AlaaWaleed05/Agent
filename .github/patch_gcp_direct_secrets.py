from pathlib import Path

path = Path('streamlit_app (7).py')
text = path.read_text(encoding='utf-8')

# Keep the existing Excel column-detection patch idempotent.
marker = '\n# =========================================================\n# STATUS COLUMN\n# =========================================================\n'

block = '''
# =========================================================
# EXCEL COLUMN DETECTION / IMPORT
# Only name, email, password and status are relevant.
# Other Excel columns are preserved and ignored by the importer.
# =========================================================

def _excel_text(value):
    return "" if value is None else str(value).strip()


def _excel_normalize_header(value):
    value = _excel_text(value).lower()
    value = value.replace("إ", "ا").replace("أ", "ا").replace("آ", "ا")
    value = re.sub(r"[ًٌٍَُِّْـ]", "", value)
    return re.sub(r"[^a-z0-9\\u0600-\\u06ff]+", "", value)


def _excel_header_kind(value):
    low = _excel_normalize_header(value)

    if low in {"name", "fullname", "studentname", "student", "اسم", "الاسم", "اسمالطالب"} or ("اسم" in low and "خدمة" not in low):
        return "name"

    if low in {"email", "emailaddress", "mail", "البريد", "البريدالالكتروني", "البريدالإلكتروني", "الايميل", "الإيميل", "ايميل", "إيميل"} or any(token in low for token in ("email", "mail", "بريد", "ايميل")):
        return "email"

    if low in {"password", "pass", "studentpassword", "الباسورد", "باسورد", "كلمةالمرور", "كلمهالمرور", "كلمةالسر", "كلمهالسر", "الرقمالسري"} or any(token in low for token in ("password", "pass", "باسورد", "كلمةالمرور", "كلمهالمرور", "كلمةالسر", "كلمهالسر")):
        return "password"

    if low in {"status", "applicationstatus", "حالة", "الحالة", "حالةالطلب", "حالةالطلبات", "حالةالطلبة", "حالةالطلبةالجديدة", "حالةالطلبةالجديده", "حالةالطلاب", "حالةالطلابالجديدة", "حالةالطلابالجديده"} or ("حالة" in low and "اسم" not in low and "خدمة" not in low):
        return "status"

    return None


def _find_excel_layout(ws):
    for row_idx in range(1, min(30, ws.max_row) + 1):
        found = {}
        for col_idx, cell in enumerate(ws[row_idx], start=1):
            kind = _excel_header_kind(cell.value)
            if kind and kind not in found:
                found[kind] = col_idx
        if "email" in found and "password" in found:
            return found, row_idx

    raise ValueError(
        "ملف Excel لازم يحتوي على أعمدة الإيميل والباسورد."
    )


def find_excel_columns(ws):
    cols, header_row = _find_excel_layout(ws)
    return {
        "name": cols.get("name", cols["email"]) - 1,
        "email": cols["email"] - 1,
        "password": cols["password"] - 1,
    }, header_row


def parse_excel_bytes(file_bytes):
    if not file_bytes:
        raise ValueError("ملف Excel غير موجود.")

    try:
        wb = openpyxl.load_workbook(
            io.BytesIO(file_bytes),
            data_only=False,
            read_only=False
        )
    except Exception as exc:
        raise ValueError(
            "تعذر قراءة ملف Excel. استخدمي ملف .xlsx صالح."
        ) from exc

    ws = wb.active
    cols, header_row = find_excel_columns(ws)
    records = []
    seen = set()

    for excel_row, row in enumerate(
        ws.iter_rows(min_row=header_row + 1, values_only=True),
        start=header_row + 1,
    ):
        values = list(row)

        def cell_at(key):
            idx = cols[key]
            return _excel_text(values[idx]) if idx < len(values) else ""

        email = cell_at("email")
        password = cell_at("password")
        name = cell_at("name") or email
        key = email.casefold()

        if not email or not password or key in seen:
            continue

        seen.add(key)
        records.append({
            "source_row_number": excel_row,
            "student_name": name,
            "login_identifier": email,
            "password": password,
            "original_data": {
                f"column_{i + 1}": _excel_text(value)
                for i, value in enumerate(values)
            },
        })

    if not records:
        raise ValueError(
            "مش لاقي طلاب عندهم إيميل وباسورد صالحين في ملف Excel."
        )

    return records

'''

if 'def parse_excel_bytes(file_bytes):' not in text:
    if marker not in text:
        raise SystemExit('status column marker not found')
    text = text.replace(marker, '\n' + block + marker, 1)

# Restore the Google Sheets helpers that were accidentally removed from the app.
if 'def get_saved_gsheet_link(' not in text or 'def save_gsheet_link(' not in text:
    anchor = 'def _excel_text(value):'
    if anchor not in text:
        raise SystemExit('Excel helper anchor not found')

    gsheet_block = '''def get_saved_gsheet_link(office_id):

    try:

        rows = (
            db()
            .table("data_sources")
            .select("source_url,created_at")
            .eq("office_id", office_id)
            .eq("source_type", "google_sheet")
            .not_.is_("source_url", "null")
            .order("created_at", desc=True)
            .limit(1)
            .execute()
            .data
            or []
        )

        return (
            rows[0].get("source_url")
            if rows
            else None
        )

    except Exception:
        return None


def save_gsheet_link(office_id, link):

    try:

        if not extract_sheet_id(link):
            return False, "الرابط غير صحيح!"

        (
            db()
            .table("data_sources")
            .insert({
                "office_id": office_id,
                "source_type": "google_sheet",
                "source_name": "Google Sheet",
                "source_url": link,
                "column_mapping": {},
            })
            .execute()
        )

        return True, "تم حفظ الرابط بنجاح"

    except Exception:
        safe_log("Google Sheet link save failed")
        return False, "تعذر حفظ الرابط حاليًا."


'''

    text = text.replace(anchor, gsheet_block + anchor, 1)

compile(text, 'streamlit_app (7).py', 'exec')
path.write_text(text, encoding='utf-8')
print('Ensured Excel import helpers and Google Sheets helpers exist.')
