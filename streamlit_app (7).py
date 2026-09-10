"""
Applyra - Agent
Your Smarter Support for Every Student's Application
"""

import base64
import io
import json
import os
import re
import threading
import time
import uuid
from datetime import datetime, timezone

import bcrypt
import openpyxl
import pandas as pd
import requests
import streamlit as st
from cryptography.fernet import Fernet
from PIL import Image
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

SUPABASE_STORAGE_BUCKET = os.getenv(
    "SUPABASE_STORAGE_BUCKET",
    "excel-files"
)

# Keep fallback pacing aligned with the Worker timing policy.
LOGIN_PAGE_DELAY_MIN, LOGIN_PAGE_DELAY_MAX = 0.8, 1.5
POST_LOGIN_DELAY_SECONDS = 1.0
INBOX_DELAY_MIN, INBOX_DELAY_MAX = 0.7, 1.5
STUDENT_DELAY_MIN, STUDENT_DELAY_MAX = 4, 8

TECH_FAILURE_STATUS = "تعذر فحص الطالب حاليًا"

FINAL_STATUSES = {
    "مقبول نهائي",
    "قبول نهائي",
    "تم الرفض",
    "مرفوض نهائيًا",
    "مرفوض نهائيا",
    "مرفوض",
    "خالص",
}

ADMIN_USERNAME = st.secrets.get("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD", "admin123")


# =========================================================
# SUPABASE
# =========================================================

@st.cache_resource(show_spinner=False)
def get_supabase() -> Client:
    url = st.secrets.get(
        "SUPABASE_URL",
        os.getenv("SUPABASE_URL")
    )

    key = st.secrets.get(
        "SUPABASE_SERVICE_ROLE_KEY",
        os.getenv("SUPABASE_SERVICE_ROLE_KEY")
    )

    if not url or not key:
        raise RuntimeError("Supabase configuration is missing")

    return create_client(url, key)


def db():
    return get_supabase()


def now_iso():
    return datetime.now(timezone.utc).isoformat()


def safe_log(message):
    print(f"[Applyra] {message}")


def fallback_delay(a, b):
    """
    Use the same timing ranges as worker.py
    without touching Streamlit state.
    """
    import random

    time.sleep(
        random.uniform(a, b)
    )


# =========================================================
# OFFICE
# =========================================================

def get_office_by_name(name):
    rows = (
        db()
        .table("offices")
        .select("id,name,email,status,created_at")
        .eq("name", str(name).strip())
        .limit(1)
        .execute()
        .data
        or []
    )

    return rows[0] if rows else None


def get_office_with_hash(office_id):
    rows = (
        db()
        .table("offices")
        .select(
            "id,name,email,status,password_hash,created_at"
        )
        .eq("id", office_id)
        .limit(1)
        .execute()
        .data
        or []
    )

    return rows[0] if rows else None


def get_office_by_email(email):
    rows = (
        db()
        .table("offices")
        .select("id,name,email,status")
        .eq(
            "email",
            str(email).strip().lower()
        )
        .limit(1)
        .execute()
        .data
        or []
    )

    return rows[0] if rows else None


def register_office(name, email, password):

    name = name.strip()
    email = email.strip().lower()
    password = password.strip()

    if not name or not email or not password:
        return False, "كمّلي كل البيانات الأول."

    if len(password) < 6:
        return False, "الباسورد لازم يكون 6 أحرف على الأقل."

    try:

        if get_office_by_name(name) or get_office_by_email(email):
            return False, "اسم المكتب أو الإيميل مسجل قبل كده."

        password_hash = (
            bcrypt
            .hashpw(
                password.encode(),
                bcrypt.gensalt()
            )
            .decode()
        )

        (
            db()
            .table("offices")
            .insert({
                "name": name,
                "email": email,
                "password_hash": password_hash,
                "status": "pending"
            })
            .execute()
        )

        return (
            True,
            "تم التسجيل بنجاح. الحساب في انتظار موافقة الإدارة."
        )

    except Exception:

        safe_log(
            "office registration failed"
        )

        return (
            False,
            "تعذر إنشاء الحساب حاليًا. حاولي مرة تانية."
        )


def check_login(name, password):

    try:

        office = get_office_by_name(name)

        if not office:
            return False, "اسم المكتب أو الباسورد غلط!"

        if office.get("status") == "pending":
            return False, "حسابك في انتظار موافقة الإدارة!"

        if office.get("status") != "approved":
            return False, "حسابك موقوف!"

        full = get_office_with_hash(
            office["id"]
        )

        password_hash = str(
            (full or {}).get("password_hash")
            or ""
        )

        if (
            not password_hash
            or not bcrypt.checkpw(
                password.encode(),
                password_hash.encode()
            )
        ):
            return False, "اسم المكتب أو الباسورد غلط!"

        return True, office

    except Exception:

        safe_log(
            "login check failed"
        )

        return (
            False,
            "تعذر تسجيل الدخول حاليًا. حاولي مرة تانية."
        )


def get_pending_accounts():

    try:

        return (
            db()
            .table("offices")
            .select(
                "id,name,email,status,created_at"
            )
            .eq("status", "pending")
            .order("created_at")
            .execute()
            .data
            or []
        )

    except Exception:
        return []


def set_office_status(office_id, status):

    (
        db()
        .table("offices")
        .update({
            "status": status
        })
        .eq("id", office_id)
        .execute()
    )


def log_activity(
    office_id,
    action,
    file_name="",
    details=None,
    student_id=None,
    data_source_id=None
):

    try:

        (
            db()
            .table("activity_logs")
            .insert({
                "office_id": office_id,
                "student_record_id": student_id,
                "data_source_id": data_source_id,
                "action": action,
                "file_name": file_name,
                "details": details or {},
            })
            .execute()
        )

    except Exception as exc:

        safe_log(
            f"activity log error: {exc}"
        )


# =========================================================
# GOOGLE SHEETS
# =========================================================

def get_gsheet_client():

    if gspread is None or Credentials is None:
        raise RuntimeError(
            "Google Sheets libraries unavailable"
        )

    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=[
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
        ],
    )

    return gspread.authorize(creds)


def extract_sheet_id(link):

    match = re.search(
        r"/spreadsheets/d/([a-zA-Z0-9-_]+)",
        str(link)
    )

    return match.group(1) if match else None


def extract_gid(link):

    match = re.search(
        r"[?#&]gid=(\d+)",
        str(link)
    )

    return int(match.group(1)) if match else None


def read_gsheet_rows(link):

    sheet_id = extract_sheet_id(link)

    if not sheet_id:
        raise ValueError(
            "رابط Google Sheets غير صحيح."
        )

    spreadsheet = (
        get_gsheet_client()
        .open_by_key(sheet_id)
    )

    gid = extract_gid(link)

    worksheet = (
        next(
            (
                w
                for w in spreadsheet.worksheets()
                if w.id == gid
            ),
            spreadsheet.sheet1
        )
        if gid is not None
        else spreadsheet.sheet1
    )

    return worksheet.get_all_values()


def get_saved_gsheet_link(office_id):

    try:

        rows = (
            db()
            .table("data_sources")
            .select(
                "source_url,created_at"
            )
            .eq(
                "office_id",
                office_id
            )
            .eq(
                "source_type",
                "google_sheet"
            )
            .not_.is_(
                "source_url",
                "null"
            )
            .order(
                "created_at",
                desc=True
            )
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

        return (
            True,
            "تم حفظ الرابط بنجاح"
        )

    except Exception:

        safe_log(
            "Google Sheet link save failed"
        )

        return (
            False,
            "تعذر حفظ الرابط حاليًا."
        )


# =========================================================
# EXCEL
# =========================================================

def find_excel_columns(ws):

    cols = {
        "name": None,
        "email": None,
        "password": None
    }

    header_row = None

    for row_idx, row in enumerate(
        ws.iter_rows(
            min_row=1,
            max_row=min(10, ws.max_row),
            values_only=True
        ),
        start=1
    ):

        values = [
            str(c).strip()
            if c is not None
            else ""
            for c in row
        ]

        if any(
            "يميل" in v
            or "mail" in v.lower()
            or "بريد" in v
            for v in values
        ):

            header_row = row_idx

            for i, cell in enumerate(values):

                low = cell.lower()

                if (
                    any(
                        k in cell
                        for k in [
                            "اسم",
                            "الإسم",
                            "الاسم"
                        ]
                    )
                    or "name" in low
                ):

                    cols["name"] = i

                elif (
                    any(
                        k in cell
                        for k in [
                            "يميل",
                            "بريد"
                        ]
                    )
                    or "mail" in low
                ):

                    cols["email"] = i

                elif any(
                    k in cell
                    for k in [
                        "باسورد",
                        "كلمة المرور",
                        "password",
                        "pass"
                    ]
                ):

                    cols["password"] = i

            break

    if header_row is None:
        raise ValueError(
            "مش لاقي هيدر الإكسيل."
        )

    if cols["email"] is None:
        raise ValueError(
            "مش لاقي عمود الإيميل."
        )

    if cols["password"] is None:
        raise ValueError(
            "مش لاقي عمود الباسورد."
        )

    if cols["name"] is None:
        cols["name"] = cols["email"]

    return cols, header_row


def parse_excel_bytes(file_bytes):

    wb = openpyxl.load_workbook(
        io.BytesIO(file_bytes),
        data_only=False
    )

    ws = wb.active

    cols, header_row = find_excel_columns(ws)

    records = []

    seen = set()

    for excel_row, row in enumerate(
        ws.iter_rows(
            min_row=header_row + 1,
            values_only=True
        ),
        start=header_row + 1
    ):

        values = list(row)

        email = (
            str(
                values[cols["email"]] or ""
            ).strip()
            if cols["email"] < len(values)
            else ""
        )

        password = (
            str(
                values[cols["password"]] or ""
            ).strip()
            if cols["password"] < len(values)
            else ""
        )

        name = (
            str(
                values[cols["name"]] or ""
            ).strip()
            if cols["name"] < len(values)
            else email
        )

        key = email.lower()

        if (
            not email
            or not password
            or key in seen
        ):
            continue

        seen.add(key)

        records.append({
            "source_row_number": excel_row,
            "student_name": name or email,
            "login_identifier": email,
            "password": password,
            "original_data": {
                f"column_{i+1}":
                    (
                        str(v)
                        if v is not None
                        else ""
                    )
                for i, v in enumerate(values)
            },
        })

    return records


def encrypt_password(password, key):

    if not key:
        raise RuntimeError(
            "Encryption key missing"
        )

    return (
        Fernet(key.encode())
        .encrypt(password.encode())
        .decode()
    )


# =========================================================
# SUPABASE STORAGE (EXCEL FILES)
# Replaces Google Drive as the storage backend for the
# uploaded Excel workbook. The "source_ref" / "file_path"
# value stored in data_sources / jobs is now a Supabase
# Storage object path inside SUPABASE_STORAGE_BUCKET,
# instead of a Google Drive file id. Everything else in the
# app (jobs, live updates, downloads) works exactly the same
# way it did with Drive, since these functions keep the same
# names and the same "give bytes / take bytes" contract.
# =========================================================

EXCEL_CONTENT_TYPE = (
    "application/vnd.openxmlformats-"
    "officedocument.spreadsheetml.sheet"
)


def storage_bucket():

    return (
        db()
        .storage
        .from_(SUPABASE_STORAGE_BUCKET)
    )


def upload_excel_to_storage(
    file_bytes,
    filename,
    office
):
    """
    Uploads an Excel file to Supabase Storage and returns the
    storage object path. This path is saved as source_ref /
    file_path and used later to read/update/download the file,
    exactly like the Google Drive file id used to be.
    """

    safe_name = re.sub(
        r"[^A-Za-z0-9._-]+",
        "_",
        str(filename or "students.xlsx")
    )

    storage_path = f"{uuid.uuid4().hex}_{safe_name}"

    storage_bucket().upload(
        storage_path,
        file_bytes,
        {
            "content-type": EXCEL_CONTENT_TYPE,
            "upsert": "true",
        },
    )

    return storage_path


def download_excel_from_storage(file_id):
    """
    Downloads Excel file bytes from Supabase Storage.
    'file_id' is actually the Supabase Storage object path
    returned earlier by upload_excel_to_storage.
    """

    return storage_bucket().download(str(file_id))


def replace_excel_in_storage(file_id, file_bytes):
    """
    Overwrites the Excel file already stored in Supabase
    Storage at 'file_id' (the storage object path) with new
    bytes. This is the Supabase equivalent of the old
    service.files().update(...) live-Drive-update call.
    """

    storage_bucket().upload(
        str(file_id),
        file_bytes,
        {
            "content-type": EXCEL_CONTENT_TYPE,
            "upsert": "true",
        },
    )


# =========================================================
# STATUS COLUMN
# =========================================================

def find_status_column_for_output(
    ws,
    header_row
):

    for col_idx, cell in enumerate(
        ws[header_row],
        start=1
    ):

        value = str(
            cell.value or ""
        ).strip().lower()

        if (
            value in {
                "حالة الطلب",
                "الحالة"
            }
            or (
                "حالة" in value
                and "اسم" not in value
                and "خدمة" not in value
            )
        ):

            return col_idx

    new_col = ws.max_column + 1
    ws.cell(header_row, new_col).value = "حالة الطلب"
    return new_col


# =========================================================
# OLD BULK EXCEL BUILDER
# KEPT AS REQUESTED
# =========================================================

def build_updated_excel(
    file_bytes,
    students
):

    wb = openpyxl.load_workbook(
        io.BytesIO(file_bytes),
        data_only=False
    )

    ws = wb.active

    cols, header_row = find_excel_columns(ws)

    status_col = (
        find_status_column_for_output(
            ws,
            header_row
        )
    )

    by_login = {}

    by_row = {}

    for student in students:

        login = str(
            student.get(
                "login_identifier"
            ) or ""
        ).strip().lower()

        status = str(
            student.get(
                "application_status"
            ) or ""
        ).strip()

        if login and status:
            by_login[login] = status

        if (
            student.get(
                "source_row_number"
            )
            and status
        ):
            by_row[
                int(
                    student[
                        "source_row_number"
                    ]
                )
            ] = status

    for row_idx in range(
        header_row + 1,
        ws.max_row + 1
    ):

        status = None

        email = str(
            ws.cell(
                row_idx,
                cols["email"] + 1
            ).value or ""
        ).strip().lower()

        if email:
            status = by_login.get(email)

        if status is None:
            status = by_row.get(row_idx)

        if status is not None:

            ws.cell(
                row_idx,
                status_col
            ).value = status

    output = io.BytesIO()

    wb.save(output)

    return output.getvalue()


# =========================================================
# OLD BULK GOOGLE SHEET UPDATE
# KEPT AS REQUESTED
# =========================================================

def update_google_sheet_statuses(
    source_url,
    students,
    job_id=None
):

    sheet_id = extract_sheet_id(
        source_url
    )

    if not sheet_id:
        raise RuntimeError(
            "invalid_google_sheet_url"
        )

    spreadsheet = (
        get_gsheet_client()
        .open_by_key(sheet_id)
    )

    gid = extract_gid(
        source_url
    )

    worksheet = (
        next(
            (
                w
                for w in spreadsheet.worksheets()
                if w.id == gid
            ),
            spreadsheet.sheet1
        )
        if gid is not None
        else spreadsheet.sheet1
    )

    values = worksheet.get_all_values()

    header_idx = None
    email_idx = None
    status_idx = None

    for r_idx, row in enumerate(
        values[:10]
    ):

        normalized = [
            str(v or "")
            .strip()
            .lower()
            for v in row
        ]

        if any(
            "يميل" in v
            or "mail" in v
            or "بريد" in v
            for v in normalized
        ):

            header_idx = r_idx

            for i, value in enumerate(
                normalized
            ):

                if (
                    "يميل" in value
                    or "mail" in value
                    or "بريد" in value
                ):

                    email_idx = i

                if (
                    value in {
                        "حالة الطلب",
                        "الحالة"
                    }
                    or (
                        "حالة" in value
                        and "اسم" not in value
                        and "خدمة" not in value
                    )
                ):

                    status_idx = i

            break

    if (
        header_idx is None
        or email_idx is None
        or status_idx is None
    ):

        raise RuntimeError(
            "google_sheet_columns_missing"
        )

    by_login = {
        str(
            s.get(
                "login_identifier"
            ) or ""
        ).strip().lower():
            str(
                s.get(
                    "application_status"
                ) or ""
            ).strip()
        for s in students
    }

    for row_idx in range(
        header_idx + 1,
        len(values)
    ):

        # لو المكتب سجل خروج أثناء تحديث الشيت
        # نوقف فورًا وما نكملش باقي الصفوف.
        if (
            job_id
            and not job_is_active(job_id)
        ):

            safe_log(
                f"Job {job_id} cancelled "
                f"during Google Sheet finalization."
            )

            return

        login = str(
            values[row_idx][email_idx]
            if email_idx < len(
                values[row_idx]
            )
            else ""
        ).strip().lower()

        status = by_login.get(login)

        if status:

            worksheet.update_cell(
                row_idx + 1,
                status_idx + 1,
                status
            )


# =========================================================
# OLD FINALIZATION
# KEPT BUT NO LONGER CALLED BY THE LIVE WORKER
# =========================================================

def finalize_job_output(job):

    job_id = job["id"]

    # Never start finalization if the job was cancelled.
    if not job_is_active(job_id):

        safe_log(
            f"Job {job_id} cancelled before finalization."
        )

        return None

    students = get_students_for_job(
        job_id
    )

    source_type = str(
        job.get("source_type") or ""
    )

    if source_type == "excel":

        source_ref = str(
            job.get("source_ref") or ""
        ).strip()

        if not source_ref:
            raise RuntimeError(
                "excel_source_missing"
            )

        # Check again before downloading the original file.
        if not job_is_active(job_id):
            return None

        original = download_excel_from_storage(
            source_ref
        )

        # Check again before building/uploading the final file.
        if not job_is_active(job_id):
            return None

        updated = build_updated_excel(
            original,
            students
        )

        if not job_is_active(job_id):
            return None

        final_id = upload_excel_to_storage(
            updated,
            job.get("file_name")
            or "students.xlsx",
            ""
        )

        # The job could have been cancelled while upload was happening.
        if not job_is_active(job_id):

            safe_log(
                f"Job {job_id} cancelled "
                f"after final file upload."
            )

            return final_id

        (
            db()
            .table("jobs")
            .update({
                "final_drive_file_id": final_id,
                "error": None
            })
            .eq("id", job_id)
            .eq("status", "processing")
            .execute()
        )

        return final_id

    if source_type == "google_sheet":

        source_url = str(
            job.get("source_ref") or ""
        )

        if not job_is_active(job_id):
            return None

        update_google_sheet_statuses(
            source_url,
            students,
            job_id=job_id
        )

        return None

    raise RuntimeError(
        f"unsupported_source_type:{source_type}"
    )


# =========================================================
# IMPORT STUDENTS
# =========================================================

def import_students(
    office_id,
    source_type,
    source_name,
    file_bytes=None,
    source_url=None
):

    if source_type in {
        "xlsx",
        "xls",
        "excel"
    }:

        records = parse_excel_bytes(
            file_bytes
        )

    else:

        rows = read_gsheet_rows(
            source_url
        )

        if not rows:
            raise ValueError(
                "الشيت فاضي."
            )

        wb = openpyxl.Workbook()

        ws = wb.active

        for row in rows:
            ws.append(row)

        out = io.BytesIO()

        wb.save(out)

        records = parse_excel_bytes(
            out.getvalue()
        )

    if not records:
        raise ValueError(
            "مش لاقي طلاب عندهم إيميل وباسورد صالحين."
        )

    encryption_key = st.secrets.get(
        "STUDENT_PASSWORD_ENCRYPTION_KEY",
        os.getenv(
            "STUDENT_PASSWORD_ENCRYPTION_KEY"
        )
    )

    source_type = (
        "google_sheet"
        if source_type in {
            "gsheet",
            "google_sheet"
        }
        else "excel"
    )

    file_path = None

    if source_type == "excel":

        if not file_bytes:
            raise ValueError(
                "ملف Excel غير موجود."
            )

        file_path = upload_excel_to_storage(
            file_bytes,
            source_name,
            ""
        )

    source = (
        db()
        .table("data_sources")
        .insert({
            "office_id": office_id,
            "source_type": source_type,
            "source_name": source_name,
            "source_url": source_url,
            "file_path": file_path,
            "column_mapping": {},
        })
        .execute()
        .data[0]
    )

    payload = [
        {
            "office_id": office_id,
            "data_source_id": source["id"],
            "source_row_number": r[
                "source_row_number"
            ],
            "student_name": r[
                "student_name"
            ],
            "login_identifier": r[
                "login_identifier"
            ],
            "encrypted_password":
                encrypt_password(
                    r["password"],
                    encryption_key
                ),
            "application_status": "",
            "original_data": r[
                "original_data"
            ],
            "updated_at": now_iso(),
        }
        for r in records
    ]

    (
        db()
        .table("student_records")
        .insert(payload)
        .execute()
    )

    return source, len(payload)


# =========================================================
# LIVE EXCEL UPDATE
# SAME SUPABASE STORAGE FILE
# =========================================================

def update_excel_student_status(
    source_ref,
    student,
    status
):

    if not source_ref:
        raise RuntimeError(
            "excel_source_missing"
        )

    file_bytes = download_excel_from_storage(
        source_ref
    )

    wb = openpyxl.load_workbook(
        io.BytesIO(file_bytes),
        data_only=False
    )

    ws = wb.active

    cols, header_row = find_excel_columns(
        ws
    )

    status_col = (
        find_status_column_for_output(
            ws,
            header_row
        )
    )

    source_row = student.get(
        "source_row_number"
    )

    if not source_row:

        login = str(
            student.get(
                "login_identifier"
            ) or ""
        ).strip().lower()

        for row_idx in range(
            header_row + 1,
            ws.max_row + 1
        ):

            email = str(
                ws.cell(
                    row_idx,
                    cols["email"] + 1
                ).value or ""
            ).strip().lower()

            if email == login:

                source_row = row_idx

                break

    if not source_row:
        raise RuntimeError(
            "excel_student_row_missing"
        )

    ws.cell(
        int(source_row),
        status_col
    ).value = status

    output = io.BytesIO()

    wb.save(output)

    updated_bytes = output.getvalue()

    replace_excel_in_storage(
        source_ref,
        updated_bytes
    )

    safe_log(
        f"Live Excel update: "
        f"{student.get('student_name')} "
        f"-> {status}"
    )


# =========================================================
# LIVE GOOGLE SHEET UPDATE
# SAME SHEET
# =========================================================

def update_google_sheet_student_status(
    source_url,
    student,
    status
):

    sheet_id = extract_sheet_id(
        source_url
    )

    if not sheet_id:
        raise RuntimeError(
            "invalid_google_sheet_url"
        )

    spreadsheet = (
        get_gsheet_client()
        .open_by_key(sheet_id)
    )

    gid = extract_gid(
        source_url
    )

    worksheet = (
        next(
            (
                w
                for w in spreadsheet.worksheets()
                if w.id == gid
            ),
            spreadsheet.sheet1
        )
        if gid is not None
        else spreadsheet.sheet1
    )

    values = worksheet.get_all_values()

    header_idx = None
    email_idx = None
    status_idx = None

    for r_idx, row in enumerate(
        values[:10]
    ):

        normalized = [
            str(v or "")
            .strip()
            .lower()
            for v in row
        ]

        if any(
            "يميل" in v
            or "mail" in v
            or "بريد" in v
            for v in normalized
        ):

            header_idx = r_idx

            for i, value in enumerate(
                normalized
            ):

                if (
                    "يميل" in value
                    or "mail" in value
                    or "بريد" in value
                ):
                    email_idx = i

                if (
                    value in {
                        "حالة الطلب",
                        "الحالة"
                    }
                    or (
                        "حالة" in value
                        and "اسم" not in value
                        and "خدمة" not in value
                    )
                ):
                    status_idx = i

            break

    if (
        header_idx is None
        or email_idx is None
        or status_idx is None
    ):

        raise RuntimeError(
            "google_sheet_columns_missing"
        )

    target_row = None

    login = str(
        student.get(
            "login_identifier"
        ) or ""
    ).strip().lower()

    source_row = student.get(
        "source_row_number"
    )

    if source_row:

        candidate = int(
            source_row
        )

        if (
            candidate > header_idx
            and candidate <= len(values)
        ):

            candidate_login = str(
                values[candidate - 1][email_idx]
                if email_idx < len(
                    values[candidate - 1]
                )
                else ""
            ).strip().lower()

            if candidate_login == login:
                target_row = candidate

    if target_row is None:

        for row_idx in range(
            header_idx + 1,
            len(values)
        ):

            row = values[row_idx]

            row_login = str(
                row[email_idx]
                if email_idx < len(row)
                else ""
            ).strip().lower()

            if row_login == login:

                target_row = row_idx + 1

                break

    if target_row is None:
        raise RuntimeError(
            "google_sheet_student_row_missing"
        )

    worksheet.update_cell(
        target_row,
        status_idx + 1,
        status
    )

    safe_log(
        f"Live Google Sheet update: "
        f"{student.get('student_name')} "
        f"-> {status}"
    )


# =========================================================
# LIVE SOURCE UPDATE DISPATCHER
# =========================================================

def update_source_student_status(
    job,
    student,
    status
):

    source_type = str(
        job.get("source_type") or ""
    )

    if source_type == "excel":

        update_excel_student_status(
            job.get("source_ref"),
            student,
            status
        )

    elif source_type == "google_sheet":

        update_google_sheet_student_status(
            job.get("source_ref"),
            student,
            status
        )

    else:

        raise RuntimeError(
            f"unsupported_source_type:{source_type}"
        )


# =========================================================
# JOB
# =========================================================

def create_job(
    office_id,
    source,
    file_name
):

    return (
        db()
        .table("jobs")
        .insert({
            "office_id": office_id,
            "data_source_id": source["id"],
            "source_type": source["source_type"],
            "source_ref":
                source.get("file_path")
                or source.get("source_url")
                or source["id"],
            "file_name": file_name,
            "status": "pending",
        })
        .execute()
        .data[0]
    )


def get_job(job_id):

    if not job_id:
        return None

    rows = (
        db()
        .table("jobs")
        .select("*")
        .eq("id", job_id)
        .limit(1)
        .execute()
        .data
        or []
    )

    return rows[0] if rows else None


def cancel_job(
    job_id,
    office_id
):

    if not job_id or not office_id:
        return

    try:

        (
            db()
            .table("jobs")
            .update({
                "status": "cancelled",
                "finished_at": now_iso(),
                "error": "office_logout"
            })
            .eq("id", job_id)
            .eq("office_id", office_id)
            .in_(
                "status",
                [
                    "pending",
                    "processing"
                ]
            )
            .execute()
        )

        safe_log(
            f"Job {job_id} cancelled by office logout."
        )

    except Exception as exc:

        safe_log(
            f"cancel job failed: {exc}"
        )

def cancel_all_office_jobs(office_id):

    if not office_id:
        return False

    try:

        # =====================================================
        # CANCEL ALL ACTIVE JOBS FOR THIS OFFICE
        # =====================================================

        response = (
            db()
            .table("jobs")
            .update({
                "status": "cancelled",
                "finished_at": now_iso(),
                "error": "office_logout"
            })
            .eq(
                "office_id",
                office_id
            )
            .in_(
                "status",
                [
                    "pending",
                    "processing"
                ]
            )
            .select("id,status")
            .execute()
        )

        cancelled = response.data or []

        safe_log(
            f"Office {office_id}: "
            f"cancelled {len(cancelled)} active job(s) on logout."
        )

        # =====================================================
        # VERIFY:
        # Make sure there is NO pending/processing job left.
        # =====================================================

        remaining = (
            db()
            .table("jobs")
            .select("id,status")
            .eq(
                "office_id",
                office_id
            )
            .in_(
                "status",
                [
                    "pending",
                    "processing"
                ]
            )
            .execute()
            .data
            or []
        )

        if remaining:

            safe_log(
                f"WARNING: Office {office_id} still has "
                f"{len(remaining)} active job(s) after logout."
            )

            # Retry once in case another worker/request
            # changed the job during the first update.
            (
                db()
                .table("jobs")
                .update({
                    "status": "cancelled",
                    "finished_at": now_iso(),
                    "error": "office_logout"
                })
                .eq(
                    "office_id",
                    office_id
                )
                .in_(
                    "status",
                    [
                        "pending",
                        "processing"
                    ]
                )
                .execute()
            )

            # Verify one final time.
            remaining = (
                db()
                .table("jobs")
                .select("id,status")
                .eq(
                    "office_id",
                    office_id
                )
                .in_(
                    "status",
                    [
                        "pending",
                        "processing"
                    ]
                )
                .execute()
                .data
                or []
            )

        if remaining:

            safe_log(
                f"ERROR: Could not cancel all active jobs "
                f"for office {office_id}."
            )

            return False

        safe_log(
            f"Office {office_id}: "
            f"logout cancellation verified successfully."
        )

        return True

    except Exception as exc:

        safe_log(
            f"cancel all office jobs failed: "
            f"{type(exc).__name__}: {exc}"
        )

        return False
def get_job_progress_rows(job_id):

    rows = (
        db()
        .table("job_progress")
        .select(
            "student_index,total,"
            "student_name,status,created_at"
        )
        .eq(
            "job_id",
            job_id
        )
        .order("created_at")
        .execute()
        .data
        or []
    )

    latest = {}

    for row in rows:

        key = str(
            row.get(
                "student_name"
            ) or ""
        ).strip().lower()

        if key:
            latest[key] = row

    return sorted(
        latest.values(),
        key=lambda r:
            int(
                r.get(
                    "student_index"
                ) or 0
            )
    )


def get_students_for_job(job_id):

    job = get_job(job_id)

    if not job:
        return []

    rows = (
        db()
        .table("student_records")
        .select(
            "id,student_name,"
            "login_identifier,"
            "application_status,"
            "status_updated_at,"
            "source_row_number,"
            "created_at,updated_at"
        )
        .eq(
            "office_id",
            job["office_id"]
        )
        .eq(
            "data_source_id",
            job["data_source_id"]
        )
        .order(
            "source_row_number"
        )
        .execute()
        .data
        or []
    )

    latest = {}

    for row in rows:

        key = str(
            row.get(
                "login_identifier"
            )
            or row.get(
                "student_name"
            )
            or ""
        ).strip().lower()

        if not key:
            continue

        latest[key] = row

    return list(
        latest.values()
    )


def get_students(
    office_id,
    search=""
):

    rows = (
        db()
        .table("student_records")
        .select(
            "id,student_name,"
            "login_identifier,"
            "application_status,"
            "status_updated_at,"
            "source_row_number,"
            "created_at,updated_at"
        )
        .eq(
            "office_id",
            office_id
        )
        .execute()
        .data
        or []
    )

    def stamp(v):

        try:

            return datetime.fromisoformat(
                str(v or "")
                .replace(
                    "Z",
                    "+00:00"
                )
            )

        except Exception:

            return datetime.min.replace(
                tzinfo=timezone.utc
            )

    latest = {}

    for row in rows:

        key = str(
            row.get(
                "login_identifier"
            )
            or row.get(
                "student_name"
            )
            or ""
        ).strip().lower()

        if not key:
            continue

        score = (
            stamp(
                row.get(
                    "status_updated_at"
                )
            ),
            stamp(
                row.get(
                    "updated_at"
                )
            ),
            stamp(
                row.get(
                    "created_at"
                )
            )
        )

        if (
            key not in latest
            or score > latest[key][0]
        ):

            latest[key] = (
                score,
                row
            )

    rows = sorted(
        [
            x[1]
            for x in latest.values()
        ],
        key=lambda r:
            str(
                r.get(
                    "student_name"
                ) or ""
            ).lower()
    )

    q = search.strip().lower()

    return [
        r
        for r in rows
        if (
            not q
            or q in str(
                r.get(
                    "student_name"
                ) or ""
            ).lower()
        )
    ]


# =========================================================
# API FALLBACK
# =========================================================

def _legacy_api_login(
    email,
    password
):

    session = requests.Session()

    session.headers.update({
        "accept":
            "application/json, text/plain, */*",
        "accept-language":
            "ar",
        "device":
            "CITIZEN",
        "origin":
            SITE_URL,
        "referer":
            SITE_URL + "/",
        "user-agent":
            "Mozilla/5.0",
        "content-type":
            "application/json",
    })

    response = session.post(
        f"{BASE_URL}/student/login",
        json={
            "email": email,
            "password": password
        },
        timeout=30
    )

    if response.status_code not in (
        200,
        201
    ):

        return (
            None,
            None,
            "login_failed"
        )

    body = (
        response.json()
        if response.content
        else {}
    )

    return (
        session,
        body.get("token", "")
        or response.headers.get(
            "x-csrf-token",
            ""
        ),
        None
    )


def _legacy_api_get_status(
    session,
    token
):

    filt = {
        "where": {},
        "limit": 10,
        "offset": 0,
        "order":
            "statusUpdatedAt DESC",
        "fields": [
            "serviceSlug",
            "ID",
            "createdAt",
            "statusUpdatedAt",
            "activityId",
            "activityName"
        ]
    }

    headers = (
        {"x-csrf-token": token}
        if token
        else {}
    )

    response = session.get(
        f"{BASE_URL}/dynamic_services/inbox",
        params={
            "filter":
                json.dumps(filt)
        },
        headers=headers,
        timeout=30
    )

    if response.status_code not in (
        200,
        304
    ):

        raise RuntimeError(
            f"status_http_{response.status_code}"
        )

    result = (
        response
        .json()
        .get("result")
        or []
    )

    if not result:
        return "مفيش طلبات"

    activity = (
        result[0].get(
            "activityName"
        )
        or "غير محدد"
    )

    mapping = {
        "قبول الفحص الفنى":
            "القبول المبدئي",

        "قبول الفحص الفني":
            "القبول المبدئي",

        "تم السداد":
            "تم السداد",

        "تأكيد استلام الملف وصحة و اكتمال المستندات":
            "تأكيد استلام الملف وصحة واكتمال المستندات",

        "الانتظار مراجعة الطلب":
            "بانتظار مراجعة الطلب",

        "قبول من رئيس الادارة المركزية":
            "قبول من رئيس الإدارة المركزية",
    }

    return mapping.get(
        activity,
        activity
    )


def _legacy_api_logout(
    session
):

    if session is None:
        return

    try:

        session.post(
            f"{BASE_URL}/student/logout",
            json={
                "redirectUrl":
                    SITE_URL
            },
            timeout=15
        )

    except Exception:
        pass


# =========================================================
# JOB STATUS
# =========================================================

def job_is_active(job_id):

    try:

        job = get_job(job_id)

        if not job:
            return False

        return (
            str(
                job.get(
                    "status"
                ) or ""
            )
            == "processing"
        )

    except Exception as exc:

        safe_log(
            f"job activity check failed: "
            f"{type(exc).__name__}: {exc}"
        )

        return False


def job_is_cancelled(job_id):

    try:

        job = get_job(job_id)

        if not job:
            return False

        return (
            str(
                job.get(
                    "status"
                ) or ""
            )
            == "cancelled"
        )

    except Exception as exc:

        safe_log(
            f"job cancellation check failed: "
            f"{type(exc).__name__}: {exc}"
        )

        return False


def mark_job_done_if_active(
    job_id
):

    stamp = now_iso()

    response = (
        db()
        .table("jobs")
        .update({
            "status": "done",
            "finished_at": stamp,
            "error": None
        })
        .eq(
            "id",
            job_id
        )
        .eq(
            "status",
            "processing"
        )
        .select(
            "id,status"
        )
        .execute()
    )

    rows = response.data or []

    return bool(rows)


def mark_job_failed_if_active(
    job_id,
    error
):

    stamp = now_iso()

    response = (
        db()
        .table("jobs")
        .update({
            "status": "failed",
            "finished_at": stamp,
            "error": str(error)[:1000],
        })
        .eq(
            "id",
            job_id
        )
        .eq(
            "status",
            "processing"
        )
        .select(
            "id,status"
        )
        .execute()
    )

    rows = response.data or []

    return bool(rows)


# =========================================================
# FALLBACK JOB
# =========================================================

def _claim_fallback_job(
    job_id
):

    rows = (
        db()
        .table("jobs")
        .update({
            "status":
                "processing",
            "started_at":
                now_iso(),
            "claimed_by":
                "streamlit-fallback"
        })
        .eq(
            "id",
            job_id
        )
        .eq(
            "status",
            "pending"
        )
        .select("*")
        .execute()
        .data
        or []
    )

    return rows[0] if rows else None


def _run_legacy_api_fallback(
    job
):

    """
    Run the student update job.

    IMPORTANT LIVE UPDATE RULES:

    1. Students are shuffled.
    2. Every processed student is logged out.
    3. Current student is always persisted after its operation.
    4. If office logs out while current student is being processed,
       current student is saved, then the job stops.
    5. Students not yet started are never changed.
    6. Excel / Google Sheet are updated immediately per student.
    7. No finalization pass is executed.
    """

    job_id = job["id"]

    try:

        client = db()

        key = os.getenv(
            "STUDENT_PASSWORD_ENCRYPTION_KEY"
        )

        if not key:

            # Also support Streamlit secrets.
            key = st.secrets.get(
                "STUDENT_PASSWORD_ENCRYPTION_KEY",
                ""
            )

        if not key:
            raise RuntimeError(
                "encryption_key_missing"
            )

        # =====================================================
        # INITIAL JOB CHECK
        # =====================================================

        if not job_is_active(job_id):

            safe_log(
                f"Job {job_id} is no longer active. "
                f"Worker stopped."
            )

            return

        # =====================================================
        # GET STUDENTS
        # =====================================================

        students = (
            client
            .table("student_records")
            .select("*")
            .eq(
                "office_id",
                job["office_id"]
            )
            .eq(
                "data_source_id",
                job["data_source_id"]
            )
            .order(
                "source_row_number"
            )
            .execute()
            .data
            or []
        )

        unique = {}

        for student in students:

            key_id = str(
                student.get(
                    "login_identifier"
                )
                or student.get(
                    "student_name"
                )
                or ""
            ).strip().lower()

            if key_id:
                unique[key_id] = student

        students = list(
            unique.values()
        )

        # =====================================================
        # VERY IMPORTANT:
        # KEEP RANDOM SHUFFLE
        # =====================================================

        import random

        random.shuffle(students)

        total = len(students)

        # =====================================================
        # NO STUDENTS
        # =====================================================

        if not total:

            if not job_is_active(job_id):
                return

            if mark_job_done_if_active(
                job_id
            ):

                safe_log(
                    f"Job {job_id} completed: "
                    f"no students."
                )

            return

        retry_students = []

        # =========================================================
        # FIRST PASS
        # =========================================================

        for index, student in enumerate(
            students,
            1
        ):

            # -----------------------------------------------------
            # STOP BEFORE STARTING A NEW STUDENT
            #
            # This is the important boundary.
            #
            # If logout happened during the previous student,
            # that previous student has already been saved.
            #
            # We stop HERE before touching this student.
            # -----------------------------------------------------

            if not job_is_active(
                job_id
            ):

                safe_log(
                    f"Job {job_id} stopped before "
                    f"starting student {index}."
                )

                return

            name = str(
                student.get(
                    "student_name"
                )
                or student.get(
                    "login_identifier"
                )
                or "طالب"
            ).strip()

            current = str(
                student.get(
                    "application_status"
                )
                or ""
            ).strip()

            status = (
                current
                or "لم يتم الفحص بعد"
            )

            session = None

            try:

                # -------------------------------------------------
                # FINAL STATUS
                # -------------------------------------------------

                if current in FINAL_STATUSES:

                    status = current

                else:

                    password = (
                        Fernet(
                            key.encode()
                        )
                        .decrypt(
                            str(
                                student[
                                    "encrypted_password"
                                ]
                            ).encode()
                        )
                        .decode()
                    )

                    # -------------------------------------------------
                    # Check before login
                    #
                    # If office logged out BEFORE we started this
                    # student, don't touch this student.
                    # -------------------------------------------------

                    if not job_is_active(
                        job_id
                    ):
                        return

                    fallback_delay(
                        LOGIN_PAGE_DELAY_MIN,
                        LOGIN_PAGE_DELAY_MAX
                    )

                    # -------------------------------------------------
                    # If logout happened during the preparation delay,
                    # no student operation has started yet.
                    # -------------------------------------------------

                    if not job_is_active(
                        job_id
                    ):
                        return

                    session, token, error = (
                        _legacy_api_login(
                            str(
                                student[
                                    "login_identifier"
                                ]
                            ).strip(),
                            password
                        )
                    )

                    if error:

                        status = (
                            "فشل تسجيل الدخول"
                        )

                    else:

                        time.sleep(
                            POST_LOGIN_DELAY_SECONDS
                        )

                        fallback_delay(
                            INBOX_DELAY_MIN,
                            INBOX_DELAY_MAX
                        )

                        # -------------------------------------------------
                        # IMPORTANT:
                        #
                        # DO NOT return here if logout happened.
                        #
                        # The student has already been logged in.
                        # We must finish this student's cleanup,
                        # logout, save the result, and then stop.
                        # -------------------------------------------------

                        if job_is_cancelled(
                            job_id
                        ):

                            safe_log(
                                f"Job {job_id} cancelled "
                                f"while student {index} "
                                f"was logged in."
                            )

                        else:

                            status = (
                                _legacy_api_get_status(
                                    session,
                                    token
                                )
                            )

            except Exception as exc:

                status = TECH_FAILURE_STATUS

                retry_students.append(
                    (
                        index,
                        student,
                        name
                    )
                )

                safe_log(
                    f"fallback student error "
                    f"{student.get('id')}: "
                    f"{type(exc).__name__}: {exc}"
                )

            finally:

                # =====================================================
                # ALWAYS LOGOUT CURRENT STUDENT
                # =====================================================

                _legacy_api_logout(
                    session
                )

                if session is not None:

                    fallback_delay(
                        INBOX_DELAY_MIN,
                        INBOX_DELAY_MAX
                    )

            # =========================================================
            # IMPORTANT:
            #
            # DO NOT CHECK job_is_active() HERE BEFORE SAVING.
            #
            # If office logged out during this student,
            # this student MUST still be saved.
            # =========================================================

            stamp = now_iso()

            # =========================================================
            # SAVE STUDENT RESULT IN SUPABASE
            # =========================================================

            try:

                (
                    client
                    .table("student_records")
                    .update({
                        "application_status":
                            status,
                        "status_updated_at":
                            stamp,
                        "updated_at":
                            stamp
                    })
                    .eq(
                        "id",
                        student["id"]
                    )
                    .execute()
                )

            except Exception as exc:

                safe_log(
                    f"fallback student DB persistence "
                    f"error for {name}: {exc}"
                )

            # =========================================================
            # LIVE UPDATE SOURCE
            #
            # This happens immediately for the current student.
            # No finalization later.
            # =========================================================

            try:

                update_source_student_status(
                    job,
                    student,
                    status
                )

            except Exception as exc:

                # Source update should not erase the DB result.
                safe_log(
                    f"live source update failed "
                    f"for {name}: "
                    f"{type(exc).__name__}: {exc}"
                )

            # =========================================================
            # SAVE PROGRESS
            # =========================================================

            try:

                (
                    client
                    .table("job_progress")
                    .insert({
                        "job_id":
                            job_id,
                        "student_index":
                            index,
                        "total":
                            total,
                        "student_name":
                            name,
                        "status":
                            status
                    })
                    .execute()
                )

            except Exception as exc:

                safe_log(
                    f"fallback progress persistence "
                    f"error for {name}: {exc}"
                )

            # =========================================================
            # NOW CHECK CANCELLATION
            #
            # If office logged out during current student:
            #
            # current student:
            #     DONE
            #     LOGGED OUT
            #     SAVED
            #     SOURCE UPDATED
            #
            # next student:
            #     NOT TOUCHED
            # =========================================================

            if not job_is_active(
                job_id
            ):

                safe_log(
                    f"Job {job_id} stopped after "
                    f"student {index}. "
                    f"Current student was saved."
                )

                return

            # =========================================================
            # DELAY BEFORE NEXT STUDENT
            # =========================================================

            if index < total:

                if not job_is_active(
                    job_id
                ):
                    return

                fallback_delay(
                    STUDENT_DELAY_MIN,
                    STUDENT_DELAY_MAX
                )

        # =========================================================
        # RETRY TECHNICAL FAILURES
        # =========================================================

        for retry_position, (
            index,
            student,
            name
        ) in enumerate(
            retry_students
        ):

            # -----------------------------------------------------
            # Office logged out → don't retry anything.
            # -----------------------------------------------------

            if not job_is_active(
                job_id
            ):

                safe_log(
                    f"Job {job_id} cancelled "
                    f"before retry."
                )

                return

            retry_status = (
                TECH_FAILURE_STATUS
            )

            session = None

            try:

                password = (
                    Fernet(
                        key.encode()
                    )
                    .decrypt(
                        str(
                            student[
                                "encrypted_password"
                            ]
                        ).encode()
                    )
                    .decode()
                )

                fallback_delay(
                    1.0,
                    2.0
                )

                if not job_is_active(
                    job_id
                ):
                    return

                session, token, error = (
                    _legacy_api_login(
                        str(
                            student[
                                "login_identifier"
                            ]
                        ).strip(),
                        password
                    )
                )

                if error:

                    retry_status = (
                        "فشل تسجيل الدخول"
                    )

                else:

                    time.sleep(
                        POST_LOGIN_DELAY_SECONDS
                    )

                    fallback_delay(
                        INBOX_DELAY_MIN,
                        INBOX_DELAY_MAX
                    )

                    # -------------------------------------------------
                    # If cancellation happened while retry student
                    # was already logged in, DO NOT return.
                    # We still logout and save current result.
                    # -------------------------------------------------

                    if job_is_cancelled(
                        job_id
                    ):

                        safe_log(
                            f"Job {job_id} cancelled "
                            f"while retry student "
                            f"{name} was logged in."
                        )

                    else:

                        retry_status = (
                            _legacy_api_get_status(
                                session,
                                token
                            )
                        )

            except Exception as exc:

                retry_status = (
                    TECH_FAILURE_STATUS
                )

                safe_log(
                    f"fallback retry error "
                    f"{student.get('id')}: "
                    f"{type(exc).__name__}: {exc}"
                )

            finally:

                # =====================================================
                # ALWAYS LOGOUT RETRY STUDENT
                # =====================================================

                _legacy_api_logout(
                    session
                )

                if session is not None:

                    fallback_delay(
                        INBOX_DELAY_MIN,
                        INBOX_DELAY_MAX
                    )

            # =========================================================
            # SAVE RETRY RESULT
            # EVEN IF JOB WAS CANCELLED
            # =========================================================

            stamp = now_iso()

            try:

                (
                    client
                    .table("student_records")
                    .update({
                        "application_status":
                            retry_status,
                        "status_updated_at":
                            stamp,
                        "updated_at":
                            stamp
                    })
                    .eq(
                        "id",
                        student["id"]
                    )
                    .execute()
                )

            except Exception as exc:

                safe_log(
                    f"fallback retry DB persistence "
                    f"error for {name}: {exc}"
                )

            # =========================================================
            # LIVE SOURCE UPDATE FOR RETRY
            # =========================================================

            try:

                update_source_student_status(
                    job,
                    student,
                    retry_status
                )

            except Exception as exc:

                safe_log(
                    f"live retry source update failed "
                    f"for {name}: "
                    f"{type(exc).__name__}: {exc}"
                )

            # =========================================================
            # SAVE RETRY PROGRESS
            # =========================================================

            try:

                (
                    client
                    .table("job_progress")
                    .insert({
                        "job_id":
                            job_id,
                        "student_index":
                            index,
                        "total":
                            total,
                        "student_name":
                            name,
                        "status":
                            retry_status
                    })
                    .execute()
                )

            except Exception as exc:

                safe_log(
                    f"fallback retry progress "
                    f"persistence error for {name}: {exc}"
                )

            # =========================================================
            # AFTER CURRENT RETRY IS COMPLETELY SAVED
            # STOP IF CANCELLED
            # =========================================================

            if not job_is_active(
                job_id
            ):

                safe_log(
                    f"Job {job_id} stopped after "
                    f"retry student {name}."
                )

                return

        # =========================================================
        # IMPORTANT:
        #
        # THERE IS NO FINALIZATION HERE.
        #
        # Excel / Google Sheet were already updated
        # student-by-student above.
        #
        # We only mark the job done if it is STILL processing.
        # =========================================================

        if not job_is_active(
            job_id
        ):

            safe_log(
                f"Job {job_id} cancelled "
                f"before completion."
            )

            return

        marked_done = (
            mark_job_done_if_active(
                job_id
            )
        )

        if marked_done:

            safe_log(
                f"Job {job_id} completed successfully."
            )

        else:

            # This can happen if the office cancelled
            # between the last check and the conditional update.
            current_job = get_job(
                job_id
            )

            if (
                current_job
                and str(
                    current_job.get(
                        "status"
                    ) or ""
                ) == "cancelled"
            ):

                safe_log(
                    f"Job {job_id} was cancelled "
                    f"before done update."
                )

            else:

                raise RuntimeError(
                    "job_could_not_be_marked_done"
                )

    except Exception as exc:

        safe_log(
            f"fallback failed: "
            f"{type(exc).__name__}: {exc}"
        )

        try:

            # =====================================================
            # NEVER TURN CANCELLED INTO FAILED
            # =====================================================

            current_job = get_job(
                job_id
            )

            if (
                current_job
                and str(
                    current_job.get(
                        "status"
                    ) or ""
                ) == "cancelled"
            ):

                safe_log(
                    f"Job {job_id} is cancelled. "
                    f"Keeping cancelled status."
                )

                return

            # Conditional update:
            # only processing -> failed
            marked_failed = (
                mark_job_failed_if_active(
                    job_id,
                    exc
                )
            )

            if not marked_failed:

                latest_job = get_job(
                    job_id
                )

                if (
                    latest_job
                    and str(
                        latest_job.get(
                            "status"
                        ) or ""
                    ) == "cancelled"
                ):

                    safe_log(
                        f"Job {job_id} became "
                        f"cancelled before failed update."
                    )

        except Exception as db_exc:

            safe_log(
                f"fallback job finalization failed: "
                f"{db_exc}"
            )


# =========================================================
# BACKGROUND STARTER
# =========================================================

def _background_update_job(
    job_id,
    encryption_key
):

    try:

        # Keep all Streamlit state/UI calls
        # out of the background thread.
        os.environ[
            "STUDENT_PASSWORD_ENCRYPTION_KEY"
        ] = encryption_key or ""

        deadline = (
            time.monotonic()
            + WORKER_WAIT_SECONDS
        )

        while (
            time.monotonic()
            < deadline
        ):

            job = get_job(
                job_id
            )

            if not job:
                return

            status = str(
                job.get(
                    "status"
                )
                or "pending"
            )

            # Office logged out / job cancelled.
            if status != "pending":
                return

            time.sleep(2)

        # Try to claim only if the job is still pending.
        claimed = _claim_fallback_job(
            job_id
        )

        if claimed:

            _run_legacy_api_fallback(
                claimed
            )

    except Exception as exc:

        safe_log(
            f"background update failed: "
            f"{type(exc).__name__}: {exc}"
        )

        try:

            current_job = get_job(
                job_id
            )

            # VERY IMPORTANT:
            # Never change a cancelled job into failed.
            if (
                current_job
                and str(
                    current_job.get(
                        "status"
                    ) or ""
                ) != "cancelled"
            ):

                (
                    db()
                    .table("jobs")
                    .update({
                        "status":
                            "failed",
                        "finished_at":
                            now_iso(),
                        "error":
                            str(exc)[:1000]
                    })
                    .eq(
                        "id",
                        job_id
                    )
                    .eq(
                        "status",
                        "processing"
                    )
                    .execute()
                )

        except Exception as db_exc:

            safe_log(
                f"background finalization failed: "
                f"{db_exc}"
            )


# =========================================================
# STREAMLIT ACTIONS
# =========================================================

def start_update():

    if st.session_state.get(
        "update_locked"
    ):
        return

    st.session_state.update_locked = True

    st.session_state.update_start_requested = True


def reset_session_on_logout():

    office = st.session_state.get(
        "office"
    )

    # =====================================================
    # CANCEL ALL ACTIVE JOBS FOR THIS OFFICE
    # =====================================================

    if office:

        office_id = office.get("id")

        if office_id:

            cancelled = cancel_all_office_jobs(
                office_id
            )

            if not cancelled:

                st.error(
                    "تعذر إيقاف التحديث بشكل آمن. "
                    "حاولي تسجيل الخروج مرة تانية."
                )

                return

    # =====================================================
    # CLEAR THIS STREAMLIT SESSION
    # =====================================================

    st.session_state.clear()

    st.rerun()


# =========================================================
# PAGE CONFIG
# =========================================================

# =========================================================
# LOGO (base64-embedded so it works with no extra files)
# =========================================================

LOGO_BASE64 = "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"

LOGO_DATA_URI = f"data:image/png;base64,{LOGO_BASE64}"

_logo_image = Image.open(
    io.BytesIO(
        base64.b64decode(LOGO_BASE64)
    )
)

st.set_page_config(
    page_title="Applyra - Agent",
    page_icon=_logo_image,
    layout="wide",
    initial_sidebar_state="collapsed"
)


# =========================================================
# STYLE
# =========================================================

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800&display=swap');

html,body,[class*="css"],.stApp{
    font-family:'Cairo',sans-serif!important;
    direction:rtl;
    translate:no
}

.stApp,[data-testid="stAppViewContainer"]>.main{
    background:#f5f7fb;
    color:#111827
}

.block-container{
    max-width:1180px;
    padding-top:1.4rem;
    padding-bottom:3rem
}

#MainMenu,footer,header,[data-testid="stToolbar"],
[data-testid="stDecoration"],[data-testid="stStatusWidget"],
[data-testid="stSidebarNav"]{
    display:none!important
}

h1{
    font-size:32px!important;
    font-weight:800!important
}

h2{
    font-size:23px!important;
    font-weight:800!important
}

h3{
    font-size:19px!important;
    font-weight:700!important
}

.topbar{
    background:#fff;
    border:1px solid #e5e7eb;
    border-radius:16px;
    padding:13px 18px;
    display:flex;
    align-items:center;
    justify-content:space-between;
    box-shadow:0 2px 10px rgba(17,24,39,.04);
    margin-bottom:22px
}

.brand{
    display:flex;
    align-items:center;
    gap:11px
}

.brand-icon{
    width:42px;
    height:42px;
    border-radius:12px;
    background:#eff6ff;
    display:flex;
    align-items:center;
    justify-content:center;
    font-size:22px;
    overflow:hidden
}

.brand-icon img{
    width:100%;
    height:100%;
    object-fit:contain
}

.brand-title{
    font-size:18px;
    font-weight:800;
    color:#111827
}

.brand-sub{
    font-size:12px;
    color:#6b7280;
    margin-top:-2px
}

.card{
    background:#fff;
    border:1px solid #e5e7eb;
    border-radius:18px;
    padding:22px;
    box-shadow:0 3px 14px rgba(17,24,39,.045);
    margin-bottom:18px
}

.hero{
    background:linear-gradient(135deg,#fff 0%,#f8fbff 100%);
    border:1px solid #dbeafe;
    border-radius:20px;
    padding:25px 28px;
    box-shadow:0 4px 18px rgba(37,99,235,.06);
    margin-bottom:20px
}

.hero-kicker{
    color:#6b7280;
    font-size:14px;
    font-weight:600
}

.hero-title{
    color:#111827;
    font-size:28px;
    font-weight:800;
    margin-top:2px
}

.hero-title strong{
    color:#2563eb
}

.hero-desc{
    color:#6b7280;
    font-size:14px;
    margin-top:3px
}

.section-title{
    font-size:18px;
    font-weight:800;
    color:#111827;
    margin:5px 0 13px
}

.section-sub{
    color:#6b7280;
    font-size:13px;
    margin-top:-8px;
    margin-bottom:14px
}

.stTextInput label,.stFileUploader label,
.stRadio>label,.stCheckbox label{
    color:#374151!important;
    font-size:14px!important;
    font-weight:700!important
}

.stTextInput input{
    background:#fff!important;
    color:#111827!important;
    border:1px solid #d1d5db!important;
    border-radius:10px!important;
    font-size:14px!important;
    min-height:44px
}

.stTextInput input:focus{
    border-color:#2563eb!important;
    box-shadow:0 0 0 3px rgba(37,99,235,.10)!important
}

.stButton>button,.stDownloadButton>button{
    width:100%;
    min-height:44px;
    border-radius:10px!important;
    border:1px solid #2563eb!important;
    background:#2563eb!important;
    color:#fff!important;
    font-weight:700!important;
    font-size:14px!important;
    box-shadow:0 3px 8px rgba(37,99,235,.16)!important
}

.stButton>button:hover,
.stDownloadButton>button:hover{
    background:#1d4ed8!important;
    border-color:#1d4ed8!important;
    transform:translateY(-1px)
}

button[data-baseweb="tab"]{
    color:#6b7280!important;
    font-weight:700!important;
    font-size:14px!important
}

button[data-baseweb="tab"][aria-selected="true"]{
    color:#2563eb!important
}

[data-baseweb="tab-highlight"]{
    background:#2563eb!important;
    height:2px!important
}

.stRadio div[role="radiogroup"]{
    gap:10px
}

.stRadio div[role="radiogroup"] label{
    background:#fff;
    border:1px solid #e5e7eb;
    border-radius:12px;
    padding:10px 14px
}

[data-testid="stFileUploaderDropzone"]{
    background:#f8fafc!important;
    border:1.5px dashed #cbd5e1!important;
    border-radius:14px!important
}

[data-testid="InputInstructions"]{
    display:none!important
}

[data-testid="stTextInput"] button{
    display:none!important
}

.connected-box{
    background:#f0fdf4;
    border:1px solid #bbf7d0;
    border-radius:11px;
    padding:11px 14px;
    color:#166534;
    font-size:13px;
    margin-bottom:12px
}

.wait-box{
    background:#eff6ff;
    border:1px solid #bfdbfe;
    border-radius:12px;
    padding:13px 15px;
    color:#1d4ed8;
    font-weight:700
}

.success-box{
    background:#ecfdf5;
    border:1px solid #bbf7d0;
    border-radius:12px;
    padding:14px 15px;
    color:#166534;
    font-weight:800
}

.lock-box{
    background:#fff7ed;
    border:1px solid #fed7aa;
    border-radius:12px;
    padding:13px 15px;
    color:#9a3412
}

.progress-table{
    width:100%;
    border-collapse:separate;
    border-spacing:0;
    overflow:hidden;
    border:1px solid #e5e7eb;
    border-radius:12px
}

.progress-table th,
.progress-table td{
    padding:10px 12px;
    border-bottom:1px solid #eef0f4;
    font-size:13px
}

.progress-table th{
    background:#f8fafc;
    color:#6b7280;
    font-weight:800
}

.progress-table tr:last-child td{
    border-bottom:0
}

.progress-table .num{
    width:48px;
    text-align:center;
    direction:ltr
}

.progress-table .status{
    text-align:left;
    direction:rtl
}

.progress-table .name{
    text-align:right
}
</style>
""",
    unsafe_allow_html=True
)


# =========================================================
# SESSION DEFAULTS
# =========================================================

for key, default in [
    (
        "logged_in",
        False
    ),
    (
        "is_admin",
        False
    ),
    (
        "office",
        None
    ),
    (
        "update_locked",
        False
    ),
    (
        "active_job_id",
        None
    ),
    (
        "update_start_requested",
        False
    ),
    (
        "job_preparing",
        False
    ),
    (
        "pending_file_bytes",
        None
    ),
    (
        "pending_filename",
        ""
    ),
    (
        "final_file_bytes_cache",
        None
    ),
    (
        "final_file_cache_id",
        None
    ),
]:

    if key not in st.session_state:

        st.session_state[key] = default


# =========================================================
# LOGIN
# =========================================================

if (
    not st.session_state.logged_in
    and not st.session_state.is_admin
):

    st.markdown(
        f"""
<div style='text-align:center;margin:36px 0 22px'>
<img src="{LOGO_DATA_URI}" style="width:76px;height:76px;object-fit:contain">
<div style='font-size:31px;font-weight:800'>Applyra</div>
<div style='font-size:14px;color:#6b7280'>
        Your Smarter Support for Every Student's Application
</div>
</div>
""",
        unsafe_allow_html=True
    )

    left, right = st.columns(
        [1.15, 1],
        gap="large"
    )

    with left:

        st.markdown(
            """
<div style="padding:40px 20px 20px 10px;">
<div style="font-size:14px;color:#2563eb;font-weight:800;margin-bottom:8px;">
        حل بسيط لإدارة المتابعة
</div>

<div style="font-size:31px;font-weight:800;color:#111827;line-height:1.35;">
        تابع طلبات طلابك<br>من مكان واحد.
</div>

<div style="font-size:15px;color:#6b7280;line-height:1.9;margin-top:12px;max-width:480px;">
        حدّث حالات الطلبات، اربط Google Sheets، وابحث عن أي طالب بسرعة بدون متابعة يدوية.
</div>

<div style="margin-top:22px;color:#374151;font-size:14px;line-height:2.2;">
        ✓ تحديث حالات الطلاب بشكل منظم<br>
        ✓ حفظ مصدر البيانات للمكتب<br>
        ✓ بحث سريع عن حالة أي طالب
</div>
</div>
""",
            unsafe_allow_html=True
        )

    with right:

        st.markdown(
            '<div class="card">',
            unsafe_allow_html=True
        )

        tab1, tab2 = st.tabs(
            [
                "تسجيل الدخول",
                "حساب جديد"
            ]
        )

        with tab1:

            st.markdown(
                """
<div style='font-size:22px;font-weight:800;color:#111827;margin:8px 0 3px;'>
    مرحبًا بعودتك 👋
</div>

<div style='color:#6b7280;font-size:13px;margin-bottom:18px;'>
    سجل دخولك لإدارة طلبات الطلاب
</div>
""",
                unsafe_allow_html=True
            )

            username = st.text_input(
                "اسم المكتب",
                key="login_user",
                placeholder="اكتب اسم المكتب"
            )

            password = st.text_input(
                "كلمة المرور",
                type="password",
                key="login_pass",
                placeholder="اكتب كلمة المرور"
            )

            if st.button(
                "تسجيل الدخول",
                key="login_btn"
            ):

                if (
                    username
                    == ADMIN_USERNAME
                    and password
                    == ADMIN_PASSWORD
                ):

                    st.session_state.is_admin = True

                    st.rerun()

                else:

                    ok, result = check_login(
                        username,
                        password
                    )

                    if ok:
                        
                        st.session_state.logged_in = True

                        st.session_state.office = result

                        st.session_state.update_locked = False

                        st.session_state.active_job_id = None

                        st.session_state.update_start_requested = False

                        st.session_state.job_preparing = False

                        st.session_state.pending_file_bytes = None

                        st.session_state.pending_filename = ""

                        st.session_state.final_file_bytes_cache = None

                        st.session_state.final_file_cache_id = None

                        st.rerun()

                    st.error(
                        result
                    )

        with tab2:

            st.markdown(
                """
<div style='font-size:22px;font-weight:800;color:#111827;margin:8px 0 3px;'>
    إنشاء حساب
</div>

<div style='color:#6b7280;font-size:13px;margin-bottom:18px;'>
    سجل مكتبك لبدء استخدام الخدمة
</div>
""",
                unsafe_allow_html=True
            )

            new_office = st.text_input(
                "اسم المكتب",
                key="reg_office",
                placeholder="اسم المكتب"
            )

            new_email = st.text_input(
                "الإيميل",
                key="reg_email",
                placeholder="example@email.com"
            )

            new_pass = st.text_input(
                "كلمة المرور",
                type="password",
                key="reg_pass",
                placeholder="كلمة المرور"
            )

            new_pass2 = st.text_input(
                "تأكيد كلمة المرور",
                type="password",
                key="reg_pass2",
                placeholder="أعد كتابة كلمة المرور"
            )

            if st.button(
                "إنشاء الحساب",
                key="reg_btn"
            ):

                if new_pass != new_pass2:

                    st.error(
                        "كلمة المرور مش متطابقة!"
                    )

                else:

                    ok, msg = register_office(
                        new_office,
                        new_email,
                        new_pass
                    )

                    (
                        st.success
                        if ok
                        else st.error
                    )(msg)

        st.markdown(
            '</div>',
            unsafe_allow_html=True
        )

    st.stop()


# =========================================================
# ADMIN
# =========================================================

if st.session_state.is_admin:

    st.markdown(
        """
<div class='hero'>
<div class='hero-kicker'>الإدارة</div>
<div class='hero-title'>لوحة الإدارة</div>
<div class='hero-desc'>
        إدارة الحسابات الجديدة ومتابعة حالة المكاتب.
</div>
</div>
""",
        unsafe_allow_html=True
    )

    pending = get_pending_accounts()

    st.markdown(
        f"""
<div class='section-title'>
    الحسابات المعلقة
<span style='color:#2563eb'>
        ({len(pending)})
</span>
</div>
""",
        unsafe_allow_html=True
    )

    for acc in get_pending_accounts():

        c1, c2, c3, c4 = st.columns(
            [3, 3, 1, 1]
        )

        c1.write(
            acc.get("name", "")
        )

        c2.write(
            acc.get("email", "")
        )

        if c3.button(
            "قبول",
            key=f"approve_{acc['id']}"
        ):

            set_office_status(
                acc["id"],
                "approved"
            )

            st.rerun()

        if c4.button(
            "رفض",
            key=f"reject_{acc['id']}"
        ):

            set_office_status(
                acc["id"],
                "rejected"
            )

            st.rerun()

    st.markdown(
        """
<div class='section-title'
     style='margin-top:28px;'>
    كل الحسابات
</div>
""",
        unsafe_allow_html=True
    )

    try:

        all_accounts = (
            db()
            .table("offices")
            .select(
                "name,email,status,created_at"
            )
            .order(
                "created_at",
                desc=True
            )
            .execute()
            .data
            or []
        )

        if all_accounts:

            st.dataframe(
                pd.DataFrame(
                    all_accounts
                ),
                use_container_width=True,
                hide_index=True
            )

    except Exception:
        pass

    if st.button(
        "تسجيل الخروج من الإدارة",
        key="admin_logout"
    ):

        reset_session_on_logout()

    st.stop()


# =========================================================
# OFFICE
# =========================================================

office = st.session_state.office

if not office:

    st.session_state.logged_in = False

    st.rerun()

office_id = office["id"]


# =========================================================
# TOP BAR
# =========================================================

st.markdown(
    f"""
<div class="topbar">
<div class="brand">
<div class="brand-icon"><img src="{LOGO_DATA_URI}"></div>
<div>
<div class="brand-title">Applyra</div>
<div class="brand-sub">
                Your Smarter Support for Every Student's Application
</div>
</div>
</div>

<div style="font-size:13px;color:#6b7280;">
        نظام متابعة المكاتب
</div>
</div>
""",
    unsafe_allow_html=True
)


# =========================================================
# HERO
# =========================================================

_hour = datetime.now().hour

_greeting = (
    "صباح الخير"
    if _hour < 12
    else "مساء الخير"
)

st.markdown(
    f""" 
<div class="hero">
<div class="hero-kicker">
        {_greeting} 👋
</div>

<div class="hero-title">
        أهلاً بيك،
<strong>
            {office.get('name', '')}
</strong>
</div>

<div class="hero-desc">
        تابع طلبات طلابك وحدّث الحالات من مكان واحد.
</div>
</div>
""",
    unsafe_allow_html=True
)


# =========================================================
# SOURCE SELECTION
# =========================================================

st.markdown(
    '<div class="card">',
    unsafe_allow_html=True
)

st.markdown(
    """
<div class="section-title">
    مصدر بيانات الطلاب
</div>

<div class="section-sub">
    اختر الطريقة التي يحتوي بها ملف الطلاب.
</div>
""",
    unsafe_allow_html=True
)

source = st.radio(
    "",
    [
        "📂 رفع ملف Excel",
        "🔗 ربط Google Sheets"
    ],
    horizontal=True,
    label_visibility="collapsed",
    key="source_mode"
)

saved_link = get_saved_gsheet_link(
    office_id
)

file_bytes = None
filename = ""
source_url = None


# =========================================================
# EXCEL SOURCE
# =========================================================

if source == "📂 رفع ملف Excel":

    uploaded = st.file_uploader(
        "ارفع ملف Excel",
        type=[
            "xlsx",
            "xls"
        ],
        label_visibility="collapsed",
        key="excel_upload"
    )

    if uploaded:

        file_bytes = uploaded.getvalue()

        filename = uploaded.name

        st.session_state.pending_file_bytes = (
            file_bytes
        )

        st.session_state.pending_filename = (
            filename
        )

        st.success(
            f"تم اختيار الملف: {filename}"
        )

    elif st.session_state.pending_file_bytes:

        file_bytes = (
            st.session_state.pending_file_bytes
        )

        filename = (
            st.session_state.pending_filename
        )


# =========================================================
# GOOGLE SHEET SOURCE
# =========================================================

else:

    if saved_link:

        st.markdown(
            """
<div class="connected-box">
    ✓ Google Sheets متصل بالفعل لهذا المكتب
</div>
""",
            unsafe_allow_html=True
        )

        change = st.checkbox(
            "تغيير رابط الشيت",
            key="change_link"
        )

        if change:

            new_link = st.text_input(
                "رابط Google Sheets",
                key="new_link"
            )

            if st.button(
                "حفظ الرابط",
                key="save_link"
            ):

                ok, msg = save_gsheet_link(
                    office_id,
                    new_link
                )

                (
                    st.success
                    if ok
                    else st.error
                )(msg)

                if ok:
                    st.rerun()

        source_url = saved_link

    else:

        new_link = st.text_input(
            "رابط Google Sheets",
            key="first_link"
        )

        if st.button(
            "حفظ وربط الشيت",
            key="save_first_link"
        ):

            ok, msg = save_gsheet_link(
                office_id,
                new_link
            )

            (
                st.success
                if ok
                else st.error
            )(msg)

            if ok:
                st.rerun()

    if (
        source_url
        and st.button(
            "تحميل بيانات الشيت",
            key="load_sheet"
        )
    ):

        try:

            rows = read_gsheet_rows(
                source_url
            )

            wb = openpyxl.Workbook()

            ws = wb.active

            for row in rows:
                ws.append(row)

            out = io.BytesIO()

            wb.save(out)

            st.session_state.pending_file_bytes = (
                out.getvalue()
            )

            st.session_state.pending_filename = (
                "google_sheet"
            )

            st.success(
                "تم جلب البيانات. جاهزة للتحديث."
            )

        except Exception:

            st.error(
                "تعذر جلب بيانات الشيت حاليًا."
            )

    file_bytes = (
        st.session_state.pending_file_bytes
    )

    filename = (
        st.session_state.pending_filename
    )


st.markdown(
    '</div>',
    unsafe_allow_html=True
)


# =========================================================
# START UPDATE
# =========================================================

if (
    file_bytes
    and not st.session_state.update_locked
    and not st.session_state.active_job_id
):

    st.button(
        "▶ تحديث حالات الطلاب",
        key="start_update",
        on_click=start_update
    )


# =========================================================
# PREPARE NEW JOB
# =========================================================

if (
    st.session_state.update_start_requested
    and not st.session_state.active_job_id
    and not st.session_state.job_preparing
):
    running = (
    db()
    .table("jobs")
    .select("id,status,created_at,started_at")
    .eq(
        "office_id",
        office_id
    )
    .in_(
        "status",
        [
            "pending",
            "processing"
        ]
    )
    .order(
        "created_at",
        desc=True
    )
    .limit(1)
    .execute()
    .data
    or []
)


    if running:

        st.session_state.update_start_requested = False

        st.session_state.update_locked = True

        st.warning(
            "في تحديث شغال بالفعل لهذا المكتب. "
            "استني لحد ما يخلص."
        )

    else:

        st.session_state.job_preparing = True

        try:

            if source == "🔗 ربط Google Sheets":

                source_type = "google_sheet"

                source_name = "Google Sheet"

            else:

                source_type = "excel"

                source_name = (
                    filename
                    or "students.xlsx"
                )

            src, _ = import_students(
                office_id,
                source_type,
                source_name,
                file_bytes=file_bytes,
                source_url=source_url
            )

            job = create_job(
                office_id,
                src,
                source_name
            )

            st.session_state.active_job_id = (
                job["id"]
            )

            encryption_key = st.secrets.get(
                "STUDENT_PASSWORD_ENCRYPTION_KEY",
                os.getenv(
                    "STUDENT_PASSWORD_ENCRYPTION_KEY",
                    ""
                )
            )

            threading.Thread(
                target=_background_update_job,
                args=(
                    job["id"],
                    encryption_key
                ),
                daemon=True
            ).start()

        except Exception as exc:

            safe_log(
                f"job preparation failed: "
                f"{type(exc).__name__}: {exc}"
            )

            st.session_state.update_start_requested = False

            st.session_state.update_locked = False

            st.error(
                "تعذر تجهيز التحديث حاليًا. "
                "حاولي مرة تانية."
            )

        finally:

            st.session_state.job_preparing = False


# =========================================================
# PROCESSING UI
# =========================================================

@st.fragment(run_every=2)
def render_processing():

    job_id = st.session_state.get(
        "active_job_id"
    )

    if (
        not job_id
        and not st.session_state.get(
            "update_locked"
        )
    ):
        return

    st.markdown(
        '<div class="card">',
        unsafe_allow_html=True
    )

    st.markdown(
        """
<div class="section-title">
    تحديث حالات الطلاب
</div>
""",
        unsafe_allow_html=True
    )

    job = (
        get_job(job_id)
        if job_id
        else None
    )

    if not job:

        if st.session_state.get(
            "update_locked"
        ):

            st.markdown(
                """
<div class="lock-box">
    🔒 تم تشغيل تحديث بالفعل في هذه الجلسة.
    لو عايزة تبدئي تحديث جديد، سجّلي خروج وادخلي تاني.
</div>
""",
                unsafe_allow_html=True
            )

        st.markdown(
            '</div>',
            unsafe_allow_html=True
        )

        return

    status = str(
        job.get(
            "status"
        )
        or "pending"
    )

    if status == "pending":

        st.markdown(
            """
<div class="wait-box">
    ⏳ سيبدأ التحديث خلال ثواني…
</div>
""",
            unsafe_allow_html=True
        )

    elif status == "processing":

        st.markdown(
            """
<div class="wait-box">
    ▶️ بدأ التحديث.
    جاري فحص الطلاب وتحديث الحالات…
</div>
""",
            unsafe_allow_html=True
        )

    elif status == "cancelled":

        st.markdown(
            """
<div class="lock-box">
    🛑 تم إيقاف التحديث.
    الطلاب الذين تم فحصهم قبل الخروج تم حفظ نتائجهم،
    والطلاب الذين لم يبدأ فحصهم لم يتم تغييرهم.
</div>
""",
            unsafe_allow_html=True
        )

    students = get_students_for_job(
        job_id
    )

    progress_rows = get_job_progress_rows(
        job_id
    )

    progress_map = {
        str(
            r.get(
                "student_name"
            ) or ""
        ).strip().lower(): r
        for r in progress_rows
    }

    if students:

        table_rows = []

        checked = 0

        for index, student in enumerate(
            students,
            1
        ):

            name = str(
                student.get(
                    "student_name"
                )
                or student.get(
                    "login_identifier"
                )
                or "طالب"
            )

            row = progress_map.get(
                name.strip().lower()
            )

            current = (
                row.get("status")
                if row
                else str(
                    student.get(
                        "application_status"
                    )
                    or "لم يتم الفحص بعد"
                )
            )

            if row:
                checked += 1

            table_rows.append(
                (
                    index,
                    name,
                    current
                )
            )

        total = len(
            table_rows
        )

        st.progress(
            min(
                checked / max(
                    total,
                    1
                ),
                1.0
            )
        )

        st.caption(
            f"تم فحص {min(checked, total)} من {total} طالب"
        )

        if progress_rows:

            last = progress_rows[-1]

            st.info(
                f"🔄 آخر طالب تم فحصه: "
                f"**{last.get('student_name') or 'طالب'}** "
                f"— الحالة: "
                f"**{last.get('status') or ''}**"
            )

        html = [
            """
<table class="progress-table">
<thead>
<tr>
<th class="num">#</th>
<th class="name">اسم الطالب</th>
<th class="status">الحالة</th>
</tr>
</thead>
<tbody>
"""
        ]

        for (
            number,
            name,
            current
        ) in table_rows:

            html.append(
                f"""
<tr>
<td class="num">{number}</td>
<td class="name">{name}</td>
<td class="status">{current}</td>
</tr>
"""
            )

        html.append(
            """
</tbody>
</table>
"""
        )

        st.markdown(
            "".join(html),
            unsafe_allow_html=True
        )

    if status == "done":

        st.markdown(
            """
<div class="success-box"
     style="margin-top:14px">
    اكتمل التحديث 🎉
</div>
""",
            unsafe_allow_html=True
        )

        # =====================================================
        # LIVE EXCEL FILE
        #
        # There is no finalization file anymore.
        # The same source Excel file in Supabase Storage
        # was updated live.
        # So download the same source file.
        # =====================================================

        if (
            str(
                job.get(
                    "source_type"
                ) or ""
            )
            == "excel"
        ):

            live_file_id = str(
                job.get(
                    "source_ref"
                ) or ""
            ).strip()

            if live_file_id:

                try:

                    if (
                        st.session_state.get(
                            "final_file_cache_id"
                        )
                        != live_file_id
                    ):

                        st.session_state.final_file_bytes_cache = (
                            download_excel_from_storage(
                                live_file_id
                            )
                        )

                        st.session_state.final_file_cache_id = (
                            live_file_id
                        )

                    if st.session_state.get(
                        "final_file_bytes_cache"
                    ):

                        st.download_button(
                            "⬇️ تحميل ملف Excel المحدث",
                            data=(
                                st.session_state
                                .final_file_bytes_cache
                            ),
                            file_name=(
                                job.get(
                                    "file_name"
                                )
                                or "students_updated.xlsx"
                            ),
                            mime=(
                                "application/vnd.openxmlformats-"
                                "officedocument.spreadsheetml.sheet"
                            ),
                            key=(
                                f"download_final_{job_id}"
                            ),
                        )

                except Exception as exc:

                    safe_log(
                        f"live file download failed: "
                        f"{type(exc).__name__}: {exc}"
                    )

                    st.info(
                        "تعذر تجهيز الملف للتحميل حاليًا. "
                        "حاولي مرة تانية بعد قليل."
                    )

    elif status == "failed":

        # Technical details remain in Supabase logs/job.error;
        # office sees only this safe message.

        st.info(
            "تعذر إكمال التحديث حاليًا. "
            "حاولي مرة تانية بعد قليل."
        )

    elif not job_id and st.session_state.get(
        "update_locked"
    ):

        st.markdown(
            """
<div class="lock-box">
    🔒 تم تشغيل تحديث بالفعل في هذه الجلسة.
    لو عايزة تبدئي تحديث جديد، سجّلي خروج وادخلي تاني.
</div>
""",
            unsafe_allow_html=True
        )

    st.markdown(
        '</div>',
        unsafe_allow_html=True
    )


render_processing()


# =========================================================
# SEARCH
# =========================================================

st.markdown(
    '<div class="card">',
    unsafe_allow_html=True
)

st.markdown(
    """
<div class="section-title">
    البحث عن طالب
</div>

<div class="section-sub">
    اكتب اسم الطالب لمعرفة آخر حالة محفوظة.
</div>
""",
    unsafe_allow_html=True
)

search_query = st.text_input(
    "اسم الطالب",
    label_visibility="collapsed",
    key="student_search"
)

if search_query:

    found = get_students(
        office_id,
        search_query
    )

    if found:

        for student in found:

            status = (
                student.get(
                    "application_status"
                )
                or "لم يتم الفحص بعد"
            )

            st.markdown(
                f"""
<div style="
background:#f8fafc;
border:1px solid #e5e7eb;
border-radius:11px;
padding:12px;
margin:7px 0
">
    <b>
        👤 {student.get("student_name", "")}
    </b>
    <br>
    <span style="
        color:#1d4ed8;
        font-size:13px
    ">
        {status}
    </span>
</div>
""",
                unsafe_allow_html=True
            )

    else:

        st.info(
            "مفيش طالب بالاسم ده."
        )

st.markdown(
    '</div>',
    unsafe_allow_html=True
)


# =========================================================
# LOGOUT
# =========================================================

# Logout is intentionally at the bottom.
# Logging out cancels the active job,
# but the worker saves the CURRENT student
# before stopping.

if st.button(
    "تسجيل الخروج",
    key="logout_main"
):

    reset_session_on_logout()
