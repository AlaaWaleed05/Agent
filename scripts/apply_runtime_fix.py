from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[1]
WORKER = ROOT / "worker.py"
APP = ROOT / "streamlit_app (7).py"


def replace_once(text, pattern, replacement, flags=0, label=""):
    new, count = re.subn(pattern, replacement, text, count=1, flags=flags)
    if count != 1:
        raise RuntimeError(f"Patch target not found: {label or pattern}")
    return new


# ---------------------------------------------------------------------------
# worker.py
# ---------------------------------------------------------------------------
s = WORKER.read_text(encoding="utf-8")

# One browser/tab per job. Technical errors are handled through the HTTP
# fallback instead of repeatedly destroying and recreating Chrome.
if "import requests\n" not in s:
    s = s.replace("import random\n", "import random\nimport requests\n", 1)

if 'BASE_URL = "https://apiadm.study-in-egypt.gov.eg/api"' not in s:
    s = s.replace(
        'SITE_URL = "https://admission.study-in-egypt.gov.eg"\n',
        'SITE_URL = "https://admission.study-in-egypt.gov.eg"\n'
        'BASE_URL = "https://apiadm.study-in-egypt.gov.eg/api"\n',
        1,
    )

legacy_helpers = r'''

# ============================================================
# RELIABLE HTTP FALLBACK FOR A STUDENT
# ============================================================

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

    response = session.post(
        f"{BASE_URL}/student/login",
        json={"email": email, "password": password},
        timeout=30,
    )

    if response.status_code not in (200, 201):
        return None, None, f"login_http_{response.status_code}"

    body = response.json() if response.content else {}
    token = body.get("token", "") or response.headers.get("x-csrf-token", "")
    return session, token, None


def _legacy_api_get_status(session, token):
    filt = {
        "where": {},
        "limit": 10,
        "offset": 0,
        "order": "statusUpdatedAt DESC",
        "fields": [
            "serviceSlug",
            "ID",
            "createdAt",
            "statusUpdatedAt",
            "activityId",
            "activityName",
        ],
    }

    response = session.get(
        f"{BASE_URL}/dynamic_services/inbox",
        params={"filter": json.dumps(filt, ensure_ascii=False)},
        headers={"x-csrf-token": token} if token else {},
        timeout=30,
    )

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
        session.post(
            f"{BASE_URL}/student/logout",
            json={"redirectUrl": SITE_URL},
            timeout=15,
        )
    except Exception:
        pass


def legacy_check_student_status(email, password):
    session = None
    try:
        session, token, error = _legacy_api_login(email, password)
        if error:
            return None, error
        return _legacy_api_get_status(session, token), None
    except Exception as exc:
        return None, f"{type(exc).__name__}: {exc}"
    finally:
        _legacy_api_logout(session)
'''

if "def legacy_check_student_status" not in s:
    s = s.replace(
        "# ============================================================\n# SELENIUM\n# ============================================================\n",
        legacy_helpers + "\n# ============================================================\n# SELENIUM\n# ============================================================\n",
        1,
    )

# Always refresh every student on every newly-created job. Previous status is
# history, not a reason to skip a fresh check.
s = replace_once(
    s,
    r'if current in FINAL_STATUSES:\n\n\s+status = current\n\n\s+else:\n',
    'if True:\n',
    flags=re.M,
    label="worker final-status skip",
)

# Never intentionally restart Chrome between students.
s = s.replace("needs_browser_restart = True", "needs_browser_restart = False")

# If Selenium reports a technical problem, use the already-proven API path for
# this student before falling back to the old retry pass.
old = '''            if technical_error:\n\n                status = (\n                    TECH_FAILURE_STATUS\n                )\n\n                retry_students.append(\n                    (\n                        index,\n                        student,\n                        name\n                    )\n                )\n\n                print(\n                    f"⚠️ Technical error "\n                    f"for {name}: "\n                    f"{technical_error}"\n                )\n\n                # IMPORTANT:\n                # DO NOT restart Chrome yet.\n                #\n                # The student MUST first be saved everywhere.\n\n                needs_browser_restart = False\n'''
new = '''            if technical_error:\n\n                # Selenium can fail because the browser/page is temporarily\n                # unavailable. Do not label the student as a failure yet:\n                # use the API path for the SAME credentials and keep the one\n                # Chrome tab alive for the rest of the job.\n                api_status, api_error = legacy_check_student_status(\n                    str(student.get("login_identifier") or "").strip(),\n                    password if "password" in locals() else decrypt_student_password(\n                        student["encrypted_password"]\n                    ),\n                )\n\n                if api_status:\n                    status = api_status\n                    technical_error = None\n                    print(\n                        f"    ✓ HTTP fallback succeeded for {name}"\n                    )\n                else:\n                    status = TECH_FAILURE_STATUS\n                    retry_students.append(\n                        (index, student, name)\n                    )\n                    print(\n                        f"⚠️ Technical error for {name}: "\n                        f"{technical_error}; API fallback: {api_error}"\n                    )\n\n                needs_browser_restart = False\n'''
if old not in s:
    raise RuntimeError("Patch target not found: Selenium technical-error block")
s = s.replace(old, new, 1)

# Output/status column should be optional in Excel: create it when missing,
# exactly like a sheet where the status column is added before writing.
old_status = '''    raise RuntimeError(\n        "status_column_missing"\n    )\n\n\ndef build_updated_excel'''
new_status = '''    new_col = ws.max_column + 1\n    ws.cell(header_row, new_col).value = "حالة الطلب"\n    return new_col\n\n\ndef build_updated_excel'''
s = s.replace(old_status, new_status, 1)

# Make Excel import/header discovery match the Google Sheet tolerance.
s = s.replace("max_row=5,\n            values_only=True", "max_row=min(10, ws.max_row),\n            values_only=True", 1)

WORKER.write_text(s, encoding="utf-8")


# ---------------------------------------------------------------------------
# streamlit_app (7).py
# ---------------------------------------------------------------------------
s = APP.read_text(encoding="utf-8")

# Excel output should not fail just because the uploaded workbook did not yet
# contain a status column. Add it automatically.
old = '''    raise RuntimeError(\n        "status_column_missing"\n    )\n\n\n# =========================================================\n# OLD BULK EXCEL BUILDER'''
new = '''    new_col = ws.max_column + 1\n    ws.cell(header_row, new_col).value = "حالة الطلب"\n    return new_col\n\n\n# =========================================================\n# OLD BULK EXCEL BUILDER'''
if old not in s:
    raise RuntimeError("Patch target not found: Streamlit status column")
s = s.replace(old, new, 1)

# Import the same broad header range used by the worker/Google Sheet path.
s = s.replace("max_row=5,\n            values_only=True", "max_row=min(10, ws.max_row),\n            values_only=True", 1)

APP.write_text(s, encoding="utf-8")

# Syntax validation is part of the patch itself.
compile(WORKER.read_text(encoding="utf-8"), str(WORKER), "exec")
compile(APP.read_text(encoding="utf-8"), str(APP), "exec")
print("Aivora runtime fix applied and both Python files compile.")
