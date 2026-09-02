"""Aivora Selenium Worker.

Claims pending jobs from Supabase, checks student application status with Selenium,
and writes one live progress row per student. Technical failures are isolated per
student and the job is always finalized instead of being left in processing.
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
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from supabase import Client, create_client

SITE_URL = "https://admission.study-in-egypt.gov.eg"
LOGIN_URL = f"{SITE_URL}/login"
INBOX_URL = f"{SITE_URL}/inbox"
WAIT_TIME = 20
JOB_POLL_INTERVAL_SECONDS = 10
STUDENT_DELAY_MIN, STUDENT_DELAY_MAX = 4, 8
FINAL_STATUSES = {
    "مقبول نهائي", "قبول نهائي", "تم الرفض", "مرفوض نهائيًا",
    "مرفوض نهائيا", "مرفوض", "خالص",
}
TECH_FAILURE_STATUS = "تعذر فحص الطالب حاليًا"
WORKER_ID = f"{socket.gethostname()}-{uuid.uuid4().hex[:6]}"
CONFIG_FILE = Path(__file__).with_name("service_account.json")


def get_config():
    if not CONFIG_FILE.exists():
        return {}
    try:
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


CONFIG = get_config()


def get_setting(name, required=True):
    value = os.environ.get(name) or CONFIG.get(name)
    if required and not value:
        raise RuntimeError(f"{name} missing")
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
    # This is the actual RPC created for Aivora's atomic queue.
    response = db.rpc("claim_next_pending_job", {"p_worker_id": WORKER_ID}).execute()
    data = response.data
    if isinstance(data, dict):
        return data if data.get("id") else None
    if isinstance(data, list):
        for row in data:
            if isinstance(row, dict) and row.get("id"):
                return row
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
    db.table("job_progress").insert({
        "job_id": job_id,
        "student_index": index,
        "total": total,
        "student_name": str(student_name),
        "status": str(status),
    }).execute()


def get_office(office_id):
    rows = db.table("offices").select("id,name,email,status").eq("id", office_id).limit(1).execute().data or []
    return rows[0] if rows else None


def get_students_for_job(job):
    rows = db.table("student_records").select("*").eq("office_id", job["office_id"]).eq("data_source_id", job["data_source_id"]).order("source_row_number").execute().data or []
    latest = {}
    for row in rows:
        key = str(row.get("login_identifier") or row.get("student_name") or "").strip().lower()
        if key:
            latest[key] = row
    return list(latest.values())


def update_student_status(student_id, status):
    stamp = now_iso()
    db.table("student_records").update({
        "application_status": str(status),
        "status_updated_at": stamp,
        "updated_at": stamp,
    }).eq("id", student_id).execute()


def log_activity(office_id, action, file_name="", details=None, student_id=None, data_source_id=None):
    try:
        db.table("activity_logs").insert({
            "office_id": office_id,
            "student_record_id": student_id,
            "data_source_id": data_source_id,
            "action": action,
            "file_name": file_name,
            "details": details or {},
        }).execute()
    except Exception as exc:
        print(f"Activity log error: {exc}")


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
    changes = []
    for item in processed:
        name = str(item.get("name") or "").strip()
        status = str(item.get("status") or "").strip()
        if name and previous.get(name) != status:
            changes.append((name, status))
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
        print(f"    {msg} ({seconds:.1f}s)...")
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
    try:
        driver.delete_all_cookies()
    except Exception:
        pass
    try:
        driver.execute_script("window.localStorage.clear(); window.sessionStorage.clear();")
    except Exception:
        pass


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
            try:
                driver.find_element(By.CSS_SELECTOR, "input[type='email'], input[name='email']")
            except Exception:
                return True, False, None
            if "login" not in driver.current_url:
                return True, False, None
        return False, False, "login_failed"
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
            if not cells:
                continue
            status = cells[status_index].text.strip() if status_index is not None and status_index < len(cells) else ""
            if status:
                statuses.append(status)
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
    except Exception as exc:
        print(f"Logout warning: {exc}")
    finally:
        clear_session(driver)


def safe_quit(driver):
    if driver is not None:
        try:
            driver.quit()
        except Exception:
            pass


def process_job(job):
    job_id = str(job["id"])
    office = get_office(job.get("office_id"))
    driver = None
    processed = []
    try:
        students = get_students_for_job(job)
        random.shuffle(students)
        total = len(students)
        if not total:
            set_job_status(job_id, "done")
            return

        previous = {str(s.get("student_name") or "").strip(): str(s.get("application_status") or "").strip() for s in students if str(s.get("student_name") or "").strip()}
        print(f"=== Starting job {job_id}: {total} students ===")

        # Browser setup is inside the protected job block.
        driver = setup_browser()

        for index, student in enumerate(students, 1):
            name = str(student.get("student_name") or student.get("login_identifier") or "طالب").strip()
            current = str(student.get("application_status") or "").strip()
            status = current or "لم يتم الفحص بعد"
            technical_error = None

            try:
                if current in FINAL_STATUSES:
                    # Final applications are not logged into again, but they are
                    # still written to job_progress so the office sees every student.
                    status = current
                else:
                    password = decrypt_student_password(student["encrypted_password"])
                    ok, technical, error = selenium_login(driver, str(student["login_identifier"]).strip(), password)
                    if not ok and not technical:
                        status = "فشل تسجيل الدخول"
                    elif not ok:
                        technical_error = error
                    else:
                        ok2, error2 = selenium_go_to_inbox(driver)
                        if not ok2:
                            technical_error = error2
                        else:
                            status_text, technical2, error3 = selenium_get_status(driver)
                            if technical2:
                                technical_error = error3
                            else:
                                status = status_text
            except Exception as exc:
                technical_error = str(exc)

            if technical_error:
                status = TECH_FAILURE_STATUS
                print(f"⚠️ Technical error for {name}: {technical_error}")
                # Restart Chrome for the next student. A single student's failure
                # must not stop the rest of the queue when restart succeeds.
                safe_quit(driver)
                driver = None
                try:
                    driver = setup_browser()
                except Exception as restart_exc:
                    print(f"❌ Chrome restart failed after {name}: {restart_exc}")

            try:
                update_student_status(student["id"], status)
            except Exception as exc:
                print(f"❌ Could not update student record {name}: {exc}")

            try:
                append_progress(job_id, index, total, name, status)
            except Exception as exc:
                print(f"❌ Could not append progress for {name}: {exc}")

            processed.append({"name": name, "status": status})
            print(f"    {index}/{total} | {name} | {status}")

            # If Chrome could not be restarted, record the remaining students as
            # failed progress instead of leaving the job hanging.
            if driver is None and index < total:
                for remaining_index, remaining in enumerate(students[index:], index + 1):
                    remaining_name = str(remaining.get("student_name") or remaining.get("login_identifier") or "طالب").strip()
                    remaining_status = TECH_FAILURE_STATUS
                    try:
                        update_student_status(remaining["id"], remaining_status)
                    except Exception as exc:
                        print(f"❌ Remaining student update error: {exc}")
                    try:
                        append_progress(job_id, remaining_index, total, remaining_name, remaining_status)
                    except Exception as exc:
                        print(f"❌ Remaining progress error: {exc}")
                    processed.append({"name": remaining_name, "status": remaining_status})
                break

            if index < total:
                human_delay(STUDENT_DELAY_MIN, STUDENT_DELAY_MAX, "Pause before next student")

        try:
            notify_office_status_changes(office, previous, processed)
        except Exception as exc:
            print(f"Status email warning: {exc}")
        log_activity(job.get("office_id"), "اكتمل تحديث حالات الطلاب", job.get("file_name") or "", {"job_id": job_id, "students_processed": len(processed), "worker_id": WORKER_ID}, data_source_id=job.get("data_source_id"))
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
