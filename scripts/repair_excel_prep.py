from pathlib import Path
import re
import subprocess
import traceback

ROOT = Path(__file__).resolve().parents[1]
APP = ROOT / "streamlit_app (7).py"
RESTORE_COMMIT = "d49dee13669cb8b60c4dc90917891c010331f098"


def restore_app():
    result = subprocess.run(
        ["git", "show", f"{RESTORE_COMMIT}:streamlit_app (7).py"],
        cwd=ROOT,
        check=True,
        capture_output=True,
    )
    APP.write_bytes(result.stdout)


def replace_once(text, old, new, label):
    if old not in text:
        raise RuntimeError(f"Patch target not found: {label}")
    return text.replace(old, new, 1)


def patch_app():
    s = APP.read_text(encoding="utf-8")

    # Make preparation failures actionable without exposing passwords or file contents.
    s = replace_once(
        s,
        "import time\nfrom datetime import datetime, timezone\n",
        "import time\nimport traceback\nfrom datetime import datetime, timezone\n",
        "traceback import",
    )

    start = s.index("def import_students(\n")
    end = s.index("\n\n# =========================================================\n# LIVE EXCEL UPDATE", start)
    old_func = s[start:end]
    new_func = '''def import_students(\n    office_id,\n    source_type,\n    source_name,\n    file_bytes=None,\n    source_url=None\n):\n\n    safe_log(\n        f"[EXCEL PREP] start source_type={source_type!r}, "\n        f"source_name={source_name!r}"\n    )\n\n    try:\n        if source_type in {"xlsx", "xls", "excel"}:\n            if not file_bytes:\n                raise ValueError("ملف Excel غير موجود.")\n            safe_log("[EXCEL PREP] parsing Excel")\n            records = parse_excel_bytes(file_bytes)\n            safe_log(f"[EXCEL PREP] parsed valid records={len(records)}")\n        else:\n            safe_log("[EXCEL PREP] reading Google Sheet")\n            rows = read_gsheet_rows(source_url)\n            if not rows:\n                raise ValueError("الشيت فاضي.")\n            wb = openpyxl.Workbook()\n            ws = wb.active\n            for row in rows:\n                ws.append(row)\n            out = io.BytesIO()\n            wb.save(out)\n            records = parse_excel_bytes(out.getvalue())\n            safe_log(f"[EXCEL PREP] parsed valid records={len(records)}")\n\n        if not records:\n            raise ValueError("مش لاقي طلاب عندهم إيميل وباسورد صالحين.")\n\n        encryption_key = st.secrets.get(\n            "STUDENT_PASSWORD_ENCRYPTION_KEY",\n            os.getenv("STUDENT_PASSWORD_ENCRYPTION_KEY")\n        )\n        if not encryption_key:\n            raise RuntimeError("Encryption key missing")\n        safe_log("[EXCEL PREP] encryption key available")\n\n        source_type = (\n            "google_sheet"\n            if source_type in {"gsheet", "google_sheet"}\n            else "excel"\n        )\n\n        file_path = None\n        if source_type == "excel":\n            safe_log("[EXCEL PREP] uploading Excel to Drive")\n            file_path = upload_to_drive(file_bytes, source_name, "")\n            if not file_path:\n                raise RuntimeError("Excel Drive upload returned no file id")\n            safe_log("[EXCEL PREP] Drive upload succeeded")\n\n        safe_log("[EXCEL PREP] inserting data source")\n        source_response = (\n            db()\n            .table("data_sources")\n            .insert({\n                "office_id": office_id,\n                "source_type": source_type,\n                "source_name": source_name,\n                "source_url": source_url,\n                "file_path": file_path,\n                "column_mapping": {},\n            })\n            .execute()\n        )\n        source_rows = source_response.data or []\n        if not source_rows:\n            raise RuntimeError("data_sources insert returned no row")\n        source = source_rows[0]\n        safe_log(f"[EXCEL PREP] data source created id={source.get('id')}")\n\n        safe_log(f"[EXCEL PREP] preparing student payload count={len(records)}")\n        payload = [\n            {\n                "office_id": office_id,\n                "data_source_id": source["id"],\n                "source_row_number": r["source_row_number"],\n                "student_name": r["student_name"],\n                "login_identifier": r["login_identifier"],\n                "encrypted_password": encrypt_password(r["password"], encryption_key),\n                "application_status": "",\n                "original_data": r["original_data"],\n                "updated_at": now_iso(),\n            }\n            for r in records\n        ]\n\n        safe_log("[EXCEL PREP] inserting student records")\n        db().table("student_records").insert(payload).execute()\n        safe_log(f"[EXCEL PREP] SUCCESS imported={len(payload)}")\n        return source, len(payload)\n\n    except Exception as exc:\n        safe_log(\n            f"[EXCEL PREP] FAILED type={type(exc).__name__}: {exc}\\n"\n            f"{traceback.format_exc()}"\n        )\n        raise\n'''
    s = s[:start] + new_func + s[end:]

    old_error = '''            st.error(\n                "تعذر تجهيز التحديث حاليًا. "\n                "حاولي مرة تانية."\n            )'''
    new_error = '''            st.error(\n                "تعذر تجهيز التحديث حاليًا. "\n                f"السبب: {type(exc).__name__}: {exc}"\n            )'''
    s = replace_once(s, old_error, new_error, "job preparation error message")

    APP.write_text(s, encoding="utf-8")


def main():
    restore_app()
    patch_app()
    compile(APP.read_text(encoding="utf-8"), str(APP), "exec")
    print("Excel preparation repair applied and app compiles.")


if __name__ == "__main__":
    main()
