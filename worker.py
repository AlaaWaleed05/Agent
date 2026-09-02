"""
Aivora Selenium Worker

Worker: claim Supabase job -> login -> read application status -> update Supabase.
The important reliability rule is that every job is finalized even when Chrome/
WebDriver cannot start or a later Selenium operation crashes.
"""

import json
import os
import random
import socket
import smtplib
import time
import traceback
import uuid
from datetime import datetime, timezone
from email.mime.text import MIMEText
from pathlib import Path

from cryptography.fernet import Fernet
from supabase import create_client, Client
from selenium import webdriver
from selenium.common.exceptions import WebDriverException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

SITE_URL = "https://admission.study-in-egypt.gov.eg"
LOGIN_URL = f"{SITE_URL}/login"
INBOX_URL = f"{SITE_URL}/inbox"
WAIT_TIME = 20
JOB_POLL_INTERVAL_SECONDS = 25
STUDENT_DELAY_MIN, STUDENT_DELAY_MAX = 4, 8
FINAL_STATUSES = {"مقبول نهائي", "قبول نهائي", "تم الرفض", "مرفوض نهائيًا", "مرفوض نهائيا", "مرفوض", "خالص"}
TECH_FAILURE_STATUS = "خطأ فني في الفحص"
WORKER_ID = f"{socket.gethostname()}-{uuid.uuid4().hex[:6]}"

CONFIG_FILE = Path(__file__).with_name("service_account.json")

def get_config():
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

CONFIG = get_config()

def get_setting(name, required=True):
    value = os.environ.get(name) or CONFIG.get(name)
    if required and not value:
        raise RuntimeError(f"{name} مش موجود في Environment Variables أو service_account.json.")
    return value

def get_supabase() -> Client:
    return create_client(get_setting("SUPABASE_URL"), get_setting("SUPABASE_SERVICE_ROLE_KEY"))

db = get_supabase()

def now_iso():
    return datetime.now(timezone.utc).isoformat()

def decrypt_student_password(value):
    key = get_setting("STUDENT_PASSWORD_ENCRYPTION_KEY")
    return Fernet(key.encode()).decrypt(str(value).encode()).decode()

def claim_next_pending_job():
    data = db.rpc("claim_next_job", {"p_worker_id": WORKER_ID}).execute().data
    if isinstance(data, dict):
        return data if data.get("id") else None
    if isinstance(data, list) and data and isinstance(data[0], dict) and data[0].get("id"):
        return data[0]
    return None

def set_job_status(job_id, status, error=None):
    payload = {"status": status}
    if status == "processing":
        payload["started_at"] = now_iso()
    if status in {"done", "failed"}:
        payload["finished_at"] = now_iso()
    if error is not None:
        payload["error"] = str(error)[:1000]
    db.table("jobs").update(payload).eq("id", job_id).execute()

def append_progress(job_id, index, total, student_name, status):
    db.table("job_progress").insert({"job_id": job_id, "student_index": index, "total": total, "student_name": str(student_name), "status": str(status)}).execute()

def get_office(office_id):
    rows = db.table("offices").select("id,name,email,status").eq("id", office_id).limit(1).execute().data or []
    return rows[0] if rows else None

def get_students_for_job(job):
    q = db.table("student_records").select("*").eq("office_id", job["office_id"])
    if job.get("data_source_id"):
        q = q.eq("data_source_id", job["data_source_id"])
    return q.order("source_row_number").execute().data or []

def update_student_status(student_id, status):
    now = now_iso()
    db.table("student_records").update({"application_status": str(status), "status_updated_at": now, "updated_at": now}).eq("id", student_id).execute()

def log_activity(office_id, action, file_name="", details=None, student_id=None, data_source_id=None):
    try:
        db.table("activity_logs").insert({"office_id": office_id, "student_record_id": student_id, "data_source_id": data_source_id, "action": action, "file_name": file_name, "details": details or {}}).execute()
    except Exception as exc:
        print(f"Activity log error: {exc}")

def get_status_changes(previous, processed):
    changes = []
    for item in processed:
        name = str(item.get("name") or "").strip()
        status = str(item.get("status") or "").strip()
        if name and previous.get(name) != status:
            changes.append((name, status))
    return changes

def send_email_notification(to_email, subject, body):
    sender = get_setting("SENDER_EMAIL", required=False)
    password = get_setting("SENDER_APP_PASSWORD", required=False)
    if not sender or not password or not to_email:
        return
    try:
        msg = MIMEText(body, "plain", "utf-8")
        msg["Subject"], msg["From"], msg["To"] = subject, sender, to_email
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, password)
            server.sendmail(sender, [to_email], msg.as_string())
    except Exception as exc:
        print(f"Email error: {exc}")

def notify_office_status_changes(office, previous, processed):
    if not office or not office.get("email"):
        return
    changes = get_status_changes(previous, processed)
    if not changes:
        return
    body = [f"تحديث حالات الطلاب - {office.get('name', '')}", ""]
    body.extend(f"{i}. {name} - {status}" for i, (name, status) in enumerate(changes[:200], 1))
    send_email_notification(office["email"], f"تحديث حالات الطلاب - {office.get('name', '')}", "\n".join(body))

def notify_developer_error(office, job_id, error):
    email = get_setting("DEVELOPER_EMAIL", required=False)
    if email:
        send_email_notification(email, "Aivora Worker job failed", f"job_id: {job_id}\noffice: {office.get('name') if office else ''}\nerror: {error}")

def human_delay(a, b, msg=""):
    seconds = random.uniform(a, b)
    if msg:
        print(f"    ⏳ {msg} ({seconds:.1f}s)...")
    time.sleep(seconds)

def human_type(element, text):
    element.clear()
    human_delay(0.3, 0.7)
    for char in str(text):
        element.send_keys(char)
        time.sleep(random.uniform(0.05, 0.15))

def setup_browser():
    options = Options()
    options.add_argument("--start-maximized")
    return webdriver.Chrome(options=options)

def clear_session(driver):
    try: driver.delete_all_cookies()
    except Exception: pass
    try: driver.execute_script("window.localStorage.clear(); window.sessionStorage.clear();")
    except Exception: pass

def selenium_login(driver, email, password):
    try:
        driver.get(LOGIN_URL)
        clear_session(driver)
        driver.get(LOGIN_URL)
        wait = WebDriverWait(driver, WAIT_TIME)
        human_delay(0.8, 1.5)
        email_field = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email'], input[name='email']")))
        human_type(email_field, email)
        password_field = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']")))
        human_type(password_field, password)
        login_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit'], input[type='submit']")))
        login_btn.click()
        for _ in range(15):
            time.sleep(1)
            try: driver.find_element(By.CSS_SELECTOR, "input[type='email'], input[name='email']")
            except Exception: return True, False, None
            if "login" not in driver.current_url: return True, False, None
        return False, False, "فشل تسجيل الدخول"
    except (TimeoutException, WebDriverException) as exc:
        return False, True, str(exc)
    except Exception as exc:
        return False, True, str(exc)

def selenium_go_to_inbox(driver):
    try:
        driver.get(INBOX_URL)
        human_delay(0.7, 1.5)
        return True, None
    except (TimeoutException, WebDriverException) as exc:
        return False, str(exc)

def selenium_get_status(driver):
    try:
        wait = WebDriverWait(driver, WAIT_TIME)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table")))
        headers = driver.find_elements(By.CSS_SELECTOR, "table thead th, table tr th")
        header_texts = [h.text.strip() for h in headers]
        status_index = next((i for i, h in enumerate(header_texts) if h in ("حالة الطلب", "الحالة") or ("حالة" in h and "اسم" not in h and "خدمة" not in h)), None)
        rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        statuses = []
        for row in rows:
            cells = row.find_elements(By.CSS_SELECTOR, "td")
            if cells:
                status = cells[status_index].text.strip() if status_index is not None and status_index < len(cells) else ""
                if status: statuses.append(status)
        return (statuses[0] if statuses else "مفيش طلبات"), False, None
    except (TimeoutException, WebDriverException) as exc:
        return "", True, str(exc)
    except Exception as exc:
        return "", True, str(exc)

def selenium_logout(driver):
    try:
        wait = WebDriverWait(driver, WAIT_TIME)
        user_menu = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "[class*='user'], [class*='profile'], [class*='avatar'], [class*='account']")))
        user_menu.click()
        logout = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'تسجيل خروج') or contains(text(), 'خروج')]")))
        logout.click()
        human_delay(0.7, 1.5)
        return True
    except Exception as exc:
        print(f"Logout warning: {exc}")
        return False
    finally:
        clear_session(driver)

def safe_quit(driver):
    if driver is not None:
        try: driver.quit()
        except Exception: pass

def process_job(job):
    job_id = str(job["id"])
    office = get_office(job.get("office_id"))
    driver = None
    processed = []
    previous = {}
    try:
        students = get_students_for_job(job)
        previous = {str(s.get("student_name") or "").strip(): str(s.get("application_status") or "").strip() for s in students if str(s.get("student_name") or "").strip()}
        pending = [s for s in students if str(s.get("application_status") or "").strip() not in FINAL_STATUSES and s.get("login_identifier") and s.get("encrypted_password")]
        if not pending:
            set_job_status(job_id, "failed", "مفيش طلاب محتاجين تحديث")
            return
        random.shuffle(pending)
        total = len(pending)

        print(f"=== Starting job {job_id} ===")
        print("🖥️ Opening Chrome...")
        # CRITICAL FIX: browser startup is inside the protected block.
        # If Chrome/WebDriver fails, the job is marked failed instead of
        # remaining forever in processing.
        driver = setup_browser()

        for idx, student in enumerate(pending, 1):
            name = str(student.get("student_name") or student.get("login_identifier") or "طالب").strip()
            status = TECH_FAILURE_STATUS
            tech_error = None
            try:
                password = decrypt_student_password(student["encrypted_password"])
                ok, technical, error = selenium_login(driver, str(student["login_identifier"]).strip(), password)
                if not ok and not technical:
                    status = "فشل تسجيل الدخول"
                elif not ok:
                    tech_error = error
                else:
                    ok2, error2 = selenium_go_to_inbox(driver)
                    if not ok2:
                        tech_error = error2
                    else:
                        status_text, technical2, error3 = selenium_get_status(driver)
                        if technical2: tech_error = error3
                        else: status = status_text
            except Exception as exc:
                tech_error = str(exc)

            if tech_error:
                status = TECH_FAILURE_STATUS
                print(f"    ⚠️ Technical error for {name}: {tech_error}")
                safe_quit(driver)
                try:
                    driver = setup_browser()
                except Exception as browser_exc:
                    # The next student can still be recorded, and the job is
                    # explicitly failed rather than left in processing.
                    print(f"    ❌ Could not restart Chrome: {browser_exc}")
                    driver = None
            else:
                print(f"    ✅ {name}: {status}")
                try: selenium_logout(driver)
                except Exception: pass

            update_student_status(student["id"], status)
            append_progress(job_id, idx, total, name, status)
            processed.append({"name": name, "status": status})

            if driver is None and idx < total:
                raise RuntimeError("Chrome/WebDriver could not be started for the remaining students.")
            if idx < total:
                human_delay(STUDENT_DELAY_MIN, STUDENT_DELAY_MAX, "Pause before next student")

        try:
            notify_office_status_changes(office, previous, processed)
        except Exception as exc:
            print(f"Status email warning: {exc}")
        log_activity(job.get("office_id"), "اكتمل تحديث حالات الطلاب عبر Worker (Selenium + Supabase)", job.get("file_name") or "", {"job_id": job_id, "students_processed": len(processed), "worker_id": WORKER_ID}, data_source_id=job.get("data_source_id"))
        set_job_status(job_id, "done")
        print(f"=== Job {job_id} finished ({len(processed)}/{total}) ===")
    except Exception as exc:
        error_text = f"{type(exc).__name__}: {exc}"
        print(f"❌ Job {job_id} failed: {error_text}")
        traceback.print_exc()
        try:
            set_job_status(job_id, "failed", error_text)
        except Exception as status_exc:
            print(f"❌ Could not mark job failed: {status_exc}")
        try:
            notify_developer_error(office, job_id, error_text)
        except Exception as notify_exc:
            print(f"Developer notification warning: {notify_exc}")
    finally:
        safe_quit(driver)

def main():
    print(f"Worker (Selenium + Supabase) running... ID: {WORKER_ID}")
    while True:
        try:
            job = claim_next_pending_job()
            if job:
                process_job(job)
            else:
                time.sleep(random.uniform(JOB_POLL_INTERVAL_SECONDS * 0.7, JOB_POLL_INTERVAL_SECONDS * 1.3))
        except KeyboardInterrupt:
            print("Worker stopped manually.")
            break
        except Exception as exc:
            print(f"❌ Unexpected worker error: {type(exc).__name__}: {exc}")
            traceback.print_exc()
            time.sleep(10)

if __name__ == "__main__":
    main()
