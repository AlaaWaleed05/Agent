"""
Aivora Worker - Supabase + Selenium

Supabase-backed worker for Aivora. Selenium keeps the visible Chrome flow
against Study in Egypt; Supabase replaces the old Google Sheets queue,
student-status and progress layer.

Required environment variables:
    SUPABASE_URL
    SUPABASE_SERVICE_ROLE_KEY
    STUDENT_PASSWORD_ENCRYPTION_KEY

Optional:
    SENDER_EMAIL
    SENDER_APP_PASSWORD
    DEVELOPER_EMAIL

The service-role key and encryption key are server-side secrets only.
"""

import os
import time
import random
import socket
import uuid
import smtplib
from email.mime.text import MIMEText
from datetime import datetime, timezone

from supabase import create_client, Client
from cryptography.fernet import Fernet, InvalidToken

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
STUDENT_DELAY_MIN = 5
STUDENT_DELAY_MAX = 10
CONSECUTIVE_TECH_FAILURE_LIMIT = 5
TECH_FAILURE_STATUS = "خطأ فني في الفحص"

FINAL_STATUSES = {
    "مقبول نهائي", "قبول نهائي", "تم الرفض", "مرفوض نهائيًا",
    "مرفوض نهائيا", "مرفوض", "خالص",
}

WORKER_ID = f"{socket.gethostname()}-{uuid.uuid4().hex[:8]}"


# ==================== Supabase ====================

def get_supabase() -> Client:
    url = os.environ.get("SUPABASE_URL")
    key = os.environ.get("SUPABASE_SERVICE_ROLE_KEY")
    if not url or not key:
        raise RuntimeError(
            "Missing SUPABASE_URL or SUPABASE_SERVICE_ROLE_KEY."
        )
    return create_client(url, key)


class Database:
    def __init__(self):
        self.sb = get_supabase()

    def claim_next_pending_job(self):
        result = self.sb.rpc(
            "claim_next_job", {"p_worker_id": WORKER_ID}
        ).execute()
        rows = result.data or []
        return rows[0] if rows else None

    def set_job_status(self, job_id, status, error=None):
        payload = {"status": status, "error": error}
        if status == "processing":
            payload["started_at"] = datetime.now(timezone.utc).isoformat()
        elif status in {"done", "failed"}:
            payload["finished_at"] = datetime.now(timezone.utc).isoformat()
        self.sb.table("jobs").update(payload).eq("id", job_id).execute()

    def get_office(self, office_id):
        result = (
            self.sb.table("offices")
            .select("id,name,email")
            .eq("id", office_id)
            .single()
            .execute()
        )
        return result.data

    def get_students(self, office_id, data_source_id):
        query = (
            self.sb.table("student_records")
            .select(
                "id,office_id,data_source_id,source_row_number,student_name,"
                "login_identifier,encrypted_password,application_status,original_data"
            )
            .eq("office_id", office_id)
        )
        if data_source_id:
            query = query.eq("data_source_id", data_source_id)
        return (query.order("source_row_number").execute().data or [])

    def update_student_status(self, student_id, status):
        now = datetime.now(timezone.utc).isoformat()
        self.sb.table("student_records").update({
            "application_status": status,
            "status_updated_at": now,
            "updated_at": now,
        }).eq("id", student_id).execute()

    def append_progress(self, job_id, index, total, student_name, status):
        self.sb.table("job_progress").insert({
            "job_id": job_id,
            "student_index": index,
            "total": total,
            "student_name": student_name,
            "status": status,
        }).execute()

    def log_activity(self, office_id, data_source_id=None,
                     student_record_id=None, action="", file_name="",
                     details=None):
        self.sb.table("activity_logs").insert({
            "office_id": office_id,
            "data_source_id": data_source_id,
            "student_record_id": student_record_id,
            "action": action,
            "file_name": file_name,
            "details": details or {},
        }).execute()


# ==================== Password ====================

def decrypt_student_password(student):
    encrypted = student.get("encrypted_password")
    if not encrypted:
        raise ValueError(
            f"No encrypted password for student {student.get('student_name')}"
        )
    key = os.environ.get("STUDENT_PASSWORD_ENCRYPTION_KEY")
    if not key:
        raise RuntimeError("STUDENT_PASSWORD_ENCRYPTION_KEY is not configured")
    try:
        return Fernet(key.encode()).decrypt(str(encrypted).encode()).decode()
    except InvalidToken as exc:
        raise ValueError(
            f"Could not decrypt password for {student.get('student_name')}: invalid key/token"
        ) from exc


# ==================== Email ====================

def send_email(to_email, subject, body):
    sender = os.environ.get("SENDER_EMAIL")
    password = os.environ.get("SENDER_APP_PASSWORD")
    if not sender or not password or not to_email:
        return
    try:
        msg = MIMEText(body, "plain", "utf-8")
        msg["Subject"] = subject
        msg["From"] = sender
        msg["To"] = to_email
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, password)
            server.sendmail(sender, [to_email], msg.as_string())
    except Exception as exc:
        print(f"Email error: {exc}")


def notify_technical_failures(office_name, job_id, failures, stopped_early):
    developer_email = os.environ.get("DEVELOPER_EMAIL")
    if not developer_email or not failures:
        return
    title = "🛑 Aivora Worker - تم إيقاف المهمة" if stopped_early else "⚠️ Aivora Worker - أخطاء فنية"
    lines = [title, f"المكتب: {office_name}", f"job_id: {job_id}",
             f"عدد الأخطاء: {len(failures)}", ""]
    for i, failure in enumerate(failures[:100], 1):
        lines.append(f"{i}. {failure['name']}: {failure['error']}")
    send_email(developer_email, f"{title} - {office_name}", "\n".join(lines))


# ==================== Selenium ====================

def human_type(element, text):
    element.clear()
    time.sleep(0.7)
    for char in str(text):
        element.send_keys(char)
        time.sleep(random.uniform(0.05, 0.18))


def clear_session(driver):
    try:
        driver.delete_all_cookies()
    except Exception:
        pass
    try:
        driver.execute_script("window.localStorage.clear(); window.sessionStorage.clear();")
    except Exception:
        pass


def setup_browser():
    options = Options()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver


def restart_browser(driver):
    try:
        driver.quit()
    except Exception:
        pass
    return setup_browser()


def selenium_login(driver, email, password):
    try:
        clear_session(driver)
        driver.get(LOGIN_URL)
        wait = WebDriverWait(driver, WAIT_TIME)
        email_field = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "input[type='email'], input[name='email']")
        ))
        email_field.click()
        human_type(email_field, email)
        password_field = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
        password_field.click()
        human_type(password_field, password)
        login_btn = driver.find_element(
            By.CSS_SELECTOR, "button[type='submit'], input[type='submit']"
        )
        login_btn.click()
        for _ in range(15):
            time.sleep(1)
            try:
                driver.find_element(By.CSS_SELECTOR, "input[type='email'], input[name='email']")
                still_login = True
            except Exception:
                still_login = False
            if not still_login or "login" not in driver.current_url:
                return True, False, None
        return False, False, "فشل تسجيل الدخول"
    except (TimeoutException, WebDriverException) as exc:
        return False, True, f"خطأ فني في صفحة اللوجين: {exc}"
    except Exception as exc:
        return False, True, f"خطأ غير متوقع في اللوجين: {exc}"


def selenium_go_to_inbox(driver):
    try:
        wait = WebDriverWait(driver, WAIT_TIME)
        try:
            menu_btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(text(), 'القائمة') or contains(@class,'menu')]")
            ))
            menu_btn.click()
            time.sleep(0.8)
            my_apps = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(text(), 'طلباتي')]")
            ))
            my_apps.click()
            time.sleep(1)
        except Exception:
            driver.get(INBOX_URL)
            time.sleep(1)
        return True, None
    except (TimeoutException, WebDriverException) as exc:
        return False, f"خطأ فني في الوصول لصفحة الطلبات: {exc}"


def selenium_get_status(driver):
    try:
        wait = WebDriverWait(driver, WAIT_TIME)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table")))
        time.sleep(1.5)
        headers = driver.find_elements(By.CSS_SELECTOR, "table thead th, table tr th")
        header_texts = [h.text.strip() for h in headers]
        status_index = None
        for i, header in enumerate(header_texts):
            if header in ("حالة الطلب", "الحالة"):
                status_index = i
                break
        if status_index is None:
            for i, header in enumerate(header_texts):
                if "حالة" in header and "اسم" not in header and "خدمة" not in header:
                    status_index = i
                    break
        rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        statuses = []
        for row in rows:
            cells = row.find_elements(By.CSS_SELECTOR, "td")
            if status_index is not None and status_index < len(cells):
                value = cells[status_index].text.strip()
                if value:
                    statuses.append(value)
        if not statuses:
            return "مفيش طلبات", False, None
        return statuses[0], False, None
    except (TimeoutException, WebDriverException) as exc:
        return "", True, f"خطأ فني في جلب حالة الطلب: {exc}"
    except Exception as exc:
        return "", True, f"خطأ غير متوقع في جلب حالة الطلب: {exc}"


def selenium_logout(driver):
    try:
        wait = WebDriverWait(driver, WAIT_TIME)
        user_menu = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "[class*='user'], [class*='profile'], [class*='avatar'], [class*='account']")
        ))
        user_menu.click()
        time.sleep(1)
        logout_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[contains(text(), 'تسجيل خروج') or contains(text(), 'خروج')]")
        ))
        logout_btn.click()
        time.sleep(1)
    except Exception:
        pass
    finally:
        clear_session(driver)


# ==================== Job processing ====================

def process_job(db: Database, job):
    job_id = str(job["id"])
    office_id = job["office_id"]
    data_source_id = job.get("data_source_id")
    file_name = job.get("file_name") or "students.xlsx"
    office = db.get_office(office_id)
    office_name = office.get("name", "Unknown office")

    print(f"=== Starting Aivora job {job_id} for {office_name} ===")
    # claim_next_job already changes the row to processing atomically.

    students = db.get_students(office_id, data_source_id)
    pending = [
        s for s in students
        if str(s.get("application_status") or "").strip() not in FINAL_STATUSES
    ]

    if not pending:
        db.set_job_status(job_id, "failed", "مفيش طلاب محتاجين تحديث")
        return

    random.shuffle(pending)
    total = len(pending)
    processed_count = 0
    technical_failures = []
    consecutive_tech_failures = 0
    stopped_early = False
    driver = setup_browser()

    try:
        for index, student in enumerate(pending, 1):
            name = str(student.get("student_name") or student.get("login_identifier") or "").strip()
            email = str(student.get("login_identifier") or "").strip()
            status = None
            is_technical = False
            error_message = None
            browser_crashed = False
            login_failed = False

            print(f"👤 [{index}/{total}] {name}")
            try:
                password = decrypt_student_password(student)
                ok, is_technical, error_message = selenium_login(driver, email, password)
                if not ok and not is_technical:
                    status = "فشل تسجيل الدخول"
                    login_failed = True
                elif not ok:
                    status = TECH_FAILURE_STATUS
                else:
                    ok2, err2 = selenium_go_to_inbox(driver)
                    if not ok2:
                        is_technical = True
                        error_message = err2
                        status = TECH_FAILURE_STATUS
                    else:
                        status, is_technical, error_message = selenium_get_status(driver)
                        if is_technical:
                            status = TECH_FAILURE_STATUS
            except WebDriverException as exc:
                browser_crashed = True
                is_technical = True
                error_message = f"Chrome crashed: {exc}"
                status = TECH_FAILURE_STATUS
            except Exception as exc:
                is_technical = True
                error_message = str(exc)
                status = TECH_FAILURE_STATUS

            try:
                if browser_crashed:
                    driver = restart_browser(driver)
                elif login_failed:
                    clear_session(driver)
                else:
                    selenium_logout(driver)
            except Exception:
                clear_session(driver)

            if is_technical:
                consecutive_tech_failures += 1
                technical_failures.append({"name": name, "error": error_message or "Unknown error"})
                print(f"   ⚠️ {error_message}")
            else:
                consecutive_tech_failures = 0
                print(f"   ✅ {status}")

            db.update_student_status(student["id"], status)
            db.append_progress(job_id, index, total, name, status)
            db.log_activity(
                office_id=office_id,
                data_source_id=data_source_id,
                student_record_id=student["id"],
                action="تحديث حالة طالب عبر Worker (Selenium)",
                file_name=file_name,
                details={
                    "status": status,
                    "technical_error": bool(is_technical),
                    "worker_id": WORKER_ID,
                },
            )
            processed_count += 1

            if consecutive_tech_failures >= CONSECUTIVE_TECH_FAILURE_LIMIT:
                stopped_early = True
                print("🛑 Too many consecutive technical failures. Stopping job.")
                break

            if index < total:
                time.sleep(random.uniform(STUDENT_DELAY_MIN, STUDENT_DELAY_MAX))
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    if stopped_early:
        db.set_job_status(
            job_id,
            "failed",
            f"تم إيقاف المهمة بعد {CONSECUTIVE_TECH_FAILURE_LIMIT} أخطاء فنية متتالية",
        )
    else:
        db.set_job_status(job_id, "done")

    db.log_activity(
        office_id=office_id,
        data_source_id=data_source_id,
        action="اكتملت معالجة المهمة عبر Worker (Selenium)",
        file_name=file_name,
        details={
            "job_id": job_id,
            "processed": processed_count,
            "total": total,
            "technical_failures": len(technical_failures),
            "stopped_early": stopped_early,
            "worker_id": WORKER_ID,
        },
    )

    if technical_failures:
        notify_technical_failures(office_name, job_id, technical_failures, stopped_early)

    print(f"=== Job {job_id} finished ({processed_count}/{total}) ===")


# ==================== Main loop ====================

def main():
    print(f"Aivora Worker (Supabase + Selenium) running — {WORKER_ID}")
    print(f"Polling every ~{JOB_POLL_INTERVAL_SECONDS}s")
    db = Database()
    while True:
        try:
            job = db.claim_next_pending_job()
            if job:
                process_job(db, job)
            else:
                time.sleep(random.uniform(
                    JOB_POLL_INTERVAL_SECONDS * 0.7,
                    JOB_POLL_INTERVAL_SECONDS * 1.3,
                ))
        except KeyboardInterrupt:
            print("Worker stopped manually.")
            break
        except Exception as exc:
            print(f"Unexpected worker error: {exc}")
            time.sleep(10)


if __name__ == "__main__":
    main()
