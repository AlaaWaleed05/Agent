"""
worker.py
Aivora Selenium Worker

نفس Worker بتاع Chrome/Selenium: login -> طلباتي -> قراءة حالة الطلب -> logout.
الاختلاف الوحيد في طبقة البيانات: بدل Google Sheets، الـ jobs والطلاب والتقدم
والحالات أصبحت من Supabase، وهو نفس الـ backend الذي يستخدمه Streamlit app.

Required environment variables:
    SUPABASE_URL
    SUPABASE_SERVICE_ROLE_KEY
    STUDENT_PASSWORD_ENCRYPTION_KEY

Optional email variables:
    SENDER_EMAIL
    SENDER_APP_PASSWORD
    DEVELOPER_EMAIL

Chrome لازم يكون مثبت على الجهاز. الـ Chrome يفتح ظاهر قدامك زي النسخة القديمة.
"""

import os
import random
import socket
import smtplib
import time
import uuid
import traceback
from datetime import datetime, timezone
from email.mime.text import MIMEText

from cryptography.fernet import Fernet
from supabase import create_client, Client

from selenium import webdriver
from selenium.common.exceptions import WebDriverException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

# ==================== إعدادات ====================
SITE_URL = "https://admission.study-in-egypt.gov.eg"
LOGIN_URL = f"{SITE_URL}/login"
INBOX_URL = f"{SITE_URL}/inbox"
WAIT_TIME = 20
JOB_POLL_INTERVAL_SECONDS = 25
STUDENT_DELAY_MIN, STUDENT_DELAY_MAX = 5, 10

FINAL_STATUSES = {
    "مقبول نهائي", "قبول نهائي", "تم الرفض", "مرفوض نهائيًا",
    "مرفوض نهائيا", "مرفوض", "خالص",
}
TECH_FAILURE_STATUS = "خطأ فني في الفحص"
CONSECUTIVE_TECH_FAILURE_LIMIT = 5
WORKER_ID = f"{socket.gethostname()}-{uuid.uuid4().hex[:6]}"

# ==================== Supabase ====================
import json
import os
import json
from pathlib import Path

CONFIG_FILE = Path(__file__).with_name("service_account.json")

def get_config():
    config = {}

    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            config = json.load(f)

    return config


CONFIG = get_config()


def get_setting(name, required=True):
    value = os.environ.get(name) or CONFIG.get(name)

    if required and not value:
        raise RuntimeError(f"{name} مش موجود لا في Environment Variables ولا في Save Accounts.json.")

    return value

def get_supabase() -> Client:
    url = get_setting("SUPABASE_URL")
    key = get_setting("SUPABASE_SERVICE_ROLE_KEY")
    return create_client(url, key)


db = get_supabase()


def now_iso():
    return datetime.now(timezone.utc).isoformat()


def decrypt_student_password(encrypted_password):
    key = get_setting("STUDENT_PASSWORD_ENCRYPTION_KEY")
    return Fernet(key.encode()).decrypt(
        str(encrypted_password).encode()
    ).decode()


# ==================== Supabase helpers ====================
def claim_next_pending_job():
    result = db.rpc(
        "claim_next_job",
        {"p_worker_id": WORKER_ID},
    ).execute()

    data = result.data

    if not data:
        return None

    # لو الـRPC رجعت Dictionary لكن مفيش Job فعلية
    if isinstance(data, dict):
        if not data.get("id"):
            return None
        return data

    # لو رجعت List
    if isinstance(data, list):
        if not data:
            return None
        if isinstance(data[0], dict) and not data[0].get("id"):
            return None
        return data[0]

    return None


def set_job_status(job_id, status, error=None, final_drive_file_id=None):
    payload = {"status": status}
    if status == "processing":
        payload["started_at"] = now_iso()
    if status in {"done", "failed"}:
        payload["finished_at"] = now_iso()
    if error is not None:
        payload["error"] = str(error)[:1000]
    if final_drive_file_id:
        payload["final_drive_file_id"] = final_drive_file_id
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
    rows = (
        db.table("offices")
        .select("id,name,email,status")
        .eq("id", office_id)
        .limit(1)
        .execute()
        .data or []
    )
    return rows[0] if rows else None


def get_students_for_job(job):
    query = (
        db.table("student_records")
        .select("*")
        .eq("office_id", job["office_id"])
    )
    if job.get("data_source_id"):
        query = query.eq("data_source_id", job["data_source_id"])
    return query.order("source_row_number").execute().data or []


def get_previous_results(job):
    students = get_students_for_job(job)
    return {
        str(s.get("student_name", "")).strip(): str(s.get("application_status", "")).strip()
        for s in students
        if str(s.get("student_name", "")).strip()
    }


def update_student_status(student_id, status):
    now = now_iso()
    db.table("student_records").update({
        "application_status": str(status),
        "status_updated_at": now,
        "updated_at": now,
    }).eq("id", student_id).execute()


def log_activity(office_id, action, file_name="", details=None,
                 student_id=None, data_source_id=None):
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


def get_status_changes(previous_results, processed_results):
    changes = []
    for result in processed_results:
        name = str(result.get("name", "")).strip()
        new_status = str(result.get("status", "")).strip()
        if not name:
            continue
        old_status = previous_results.get(name)
        if old_status is None or old_status != new_status:
            changes.append({
                "name": name,
                "old_status": old_status,
                "new_status": new_status,
            })
    return changes


# ==================== Email ====================
def send_email_notification(to_email, subject, body):

    sender_email = get_setting("SENDER_EMAIL", required=False)
    sender_password = get_setting("SENDER_APP_PASSWORD", required=False)

    if not sender_email or not sender_password or not to_email:
        return

    try:
        msg = MIMEText(body, "plain", "utf-8")
        msg["Subject"] = subject
        msg["From"] = sender_email
        msg["To"] = to_email

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, [to_email], msg.as_string())

        print(f"✉️ Email sent to {to_email}")

    except Exception as exc:
        print(f"❌ Error sending email: {exc}")


def notify_office_status_changes(office, previous_results, processed_results):
    if not office:
        return
    changes = get_status_changes(previous_results, processed_results)
    if not changes or not office.get("email"):
        return
    lines = [f"تحديث حالات الطلاب - {office.get('name', '')}", ""]
    for idx, change in enumerate(changes[:200], start=1):
        lines.append(f"{idx}. {change['name']} - {change['new_status']}")
    if len(changes) > 200:
        lines.append(f"...و{len(changes) - 200} طالب/ة تاني اتغيرت حالتهم")
    lines += ["", f"إجمالي التغييرات: {len(changes)}"]
    send_email_notification(
        office["email"],
        f"تحديث حالات الطلاب - {office.get('name', '')}",
        "\n".join(lines),
    )


def notify_developer_tech_failures(office, job_id, failures, stopped_early=False):

    developer_email = get_setting("DEVELOPER_EMAIL", required=False)

    if not developer_email or not failures:
        return

    title = "⚠️ أخطاء فنية أثناء تحديث حالات الطلاب"

    if stopped_early:
        title = "🛑 " + title + " (تم إيقاف المهمة مبكرًا)"

    lines = [
        title,
        f"المكتب: {office.get('name', '') if office else ''}",
        f"job_id: {job_id}",
        "",
        f"عدد الأخطاء الفنية: {len(failures)}",
        "",
    ]

    for i, failure in enumerate(failures[:100], start=1):
        lines.append(f"{i}. {failure['name']}: {failure['error']}")

    send_email_notification(developer_email, title, "\n".join(lines))


# ==================== Selenium: نفس Chrome flow ====================
def human_type(element, text):
    element.clear()
    time.sleep(1)
    for char in str(text):
        element.send_keys(char)
        time.sleep(random.uniform(0.15, 0.4))


def slow_wait(seconds, msg=""):
    if msg:
        print(f"    ⏳ {msg} ({seconds:.1f}s)...")
    time.sleep(seconds)


def setup_browser():
    options = Options()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver


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
        slow_wait(3, "Loading login page")
        email_field = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "input[type='email'], input[name='email']")
        ))
        email_field.click()
        slow_wait(3, "Waiting before email")
        human_type(email_field, email)
        password_field = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
        password_field.click()
        slow_wait(3, "Waiting before password")
        human_type(password_field, password)
        slow_wait(2, "Waiting before click")
        login_btn = driver.find_element(By.CSS_SELECTOR, "button[type='submit'], input[type='submit']")
        login_btn.click()
        login_success = False
        for _ in range(15):
            time.sleep(1)
            try:
                driver.find_element(By.CSS_SELECTOR, "input[type='email'], input[name='email']")
                still_has_email_field = True
            except Exception:
                still_has_email_field = False
            if not still_has_email_field or "login" not in driver.current_url:
                login_success = True
                break
        if not login_success:
            return False, False, "فشل تسجيل الدخول"
        return True, False, None
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
            slow_wait(0.8, "Waiting for menu to open")
            my_apps = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(text(), 'طلباتي')]")
            ))
            my_apps.click()
            slow_wait(1, "Waiting for requests page")
        except Exception:
            driver.get(INBOX_URL)
            slow_wait(1)
        return True, None
    except (TimeoutException, WebDriverException) as exc:
        return False, f"خطأ فني في الوصول لصفحة الطلبات: {exc}"


def selenium_get_status(driver):
    try:
        wait = WebDriverWait(driver, WAIT_TIME)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table")))
        slow_wait(2)
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
        all_requests = []
        for row in rows:
            cells = row.find_elements(By.CSS_SELECTOR, "td")
            if not cells:
                continue
            status = cells[status_index].text.strip() if status_index is not None and status_index < len(cells) else ""
            if status:
                all_requests.append(status)
        if not all_requests:
            return "مفيش طلبات", False, None
        return all_requests[0], False, None
    except (TimeoutException, WebDriverException) as exc:
        return "", True, f"خطأ فني في جلب حالة الطلب: {exc}"
    except Exception as exc:
        return "", True, f"خطأ غير متوقع في جلب حالة الطلب: {exc}"


def selenium_logout(driver):
    try:
        wait = WebDriverWait(driver, WAIT_TIME)
        slow_wait(2)
        user_menu = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "[class*='user'], [class*='profile'], [class*='avatar'], [class*='account']")
        ))
        user_menu.click()
        slow_wait(2)
        logout_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[contains(text(), 'تسجيل خروج') or contains(text(), 'خروج')]")
        ))
        logout_btn.click()
        slow_wait(2)
    except Exception:
        pass
    finally:
        clear_session(driver)


def restart_browser(driver):
    try:
        driver.quit()
    except Exception:
        pass
    return setup_browser()


# ==================== Core ====================
def process_job(job):
    job_id = str(job["id"])
    office = get_office(job.get("office_id"))
    students = get_students_for_job(job)
    previous_results = {
        str(s.get("student_name", "")).strip(): str(s.get("application_status", "")).strip()
        for s in students
        if str(s.get("student_name", "")).strip()
    }

    if not students:
        set_job_status(job_id, "failed", "مفيش طلاب في الـ Supabase للـ job ده")
        return

    pending_students = []
    for student in students:
        old_status = str(student.get("application_status", "")).strip()
        if old_status in FINAL_STATUSES:
            continue
        if not student.get("login_identifier") or not student.get("encrypted_password"):
            continue
        pending_students.append(student)

    if not pending_students:
        set_job_status(job_id, "failed", "مفيش طلاب محتاجين تحديث")
        return

    processing_order = pending_students[:]
    random.shuffle(processing_order)
    total = len(processing_order)
    processed_results = []
    tech_failures = []
    tech_retry_list = []
    consecutive_tech_failures = 0
    stopped_early = False

    print(f"=== Starting job {job_id} for office {office.get('name') if office else job.get('office_id')} ===")
    print("🖥️ Opening Chrome...")
    driver = setup_browser()

    try:
        for idx, student in enumerate(processing_order):
            email = str(student["login_identifier"]).strip()
            display_name = str(student.get("student_name") or email).strip()
            try:
                password = decrypt_student_password(student["encrypted_password"])
            except Exception as exc:
                status = TECH_FAILURE_STATUS
                update_student_status(student["id"], status)
                append_progress(job_id, idx + 1, total, display_name, status)
                tech_failures.append({"name": display_name, "error": str(exc)})
                continue

            print(f"\n👤 [{idx+1}/{total}] {display_name}")
            current_status = None
            is_tech_error = False
            error_msg = None
            browser_crashed = False
            login_confirmed_failed = False

            try:
                ok, is_tech_error, error_msg = selenium_login(driver, email, password)
                if not ok and not is_tech_error:
                    current_status = "فشل تسجيل الدخول"
                    login_confirmed_failed = True
                elif not ok and is_tech_error:
                    current_status = TECH_FAILURE_STATUS
                else:
                    ok2, err2 = selenium_go_to_inbox(driver)
                    if not ok2:
                        is_tech_error = True
                        error_msg = err2
                        current_status = TECH_FAILURE_STATUS
                    else:
                        status_text, is_tech_error, error_msg = selenium_get_status(driver)
                        current_status = TECH_FAILURE_STATUS if is_tech_error else status_text
            except WebDriverException as exc:
                is_tech_error = True
                error_msg = f"الكروم اتقفل/وقع أثناء المعالجة: {exc}"
                current_status = TECH_FAILURE_STATUS
                browser_crashed = True
            finally:
                if browser_crashed:
                    driver = restart_browser(driver)
                elif login_confirmed_failed:
                    clear_session(driver)
                else:
                    selenium_logout(driver)

            if is_tech_error:
                consecutive_tech_failures += 1
                tech_failures.append({"name": display_name, "error": error_msg or "خطأ غير معروف"})
                tech_retry_list.append({"student": student})
                print(f"    ⚠️ Technical error: {error_msg}")
            else:
                consecutive_tech_failures = 0
                print(f"    ✅ Status: {current_status}")

            update_student_status(student["id"], current_status)
            append_progress(job_id, idx + 1, total, display_name, current_status)
            processed_results.append({"name": display_name, "status": str(current_status)})

            if consecutive_tech_failures >= CONSECUTIVE_TECH_FAILURE_LIMIT:
                print(f"    ⚠️ {CONSECUTIVE_TECH_FAILURE_LIMIT} technical failures in a row.")

            if idx < total - 1:
                slow_wait(random.uniform(STUDENT_DELAY_MIN, STUDENT_DELAY_MAX), "Pause before next student")

        if tech_retry_list:
            print(f"\n🔁 Retrying {len(tech_retry_list)} technical-error student(s)...")
            for retry_item in tech_retry_list:
                student = retry_item["student"]
                display_name = str(student.get("student_name") or student.get("login_identifier"))
                email = str(student["login_identifier"]).strip()
                slow_wait(random.uniform(STUDENT_DELAY_MIN, STUDENT_DELAY_MAX), "Pause before retry")
                try:
                    password = decrypt_student_password(student["encrypted_password"])
                except Exception as exc:
                    print(f"    ⚠️ Retry decrypt error: {exc}")
                    continue

                retry_status = None
                retry_is_tech_error = False
                retry_error_msg = None
                retry_login_failed = False
                retry_browser_crashed = False

                try:
                    ok, retry_is_tech_error, retry_error_msg = selenium_login(driver, email, password)
                    if not ok and not retry_is_tech_error:
                        retry_status = "فشل تسجيل الدخول"
                        retry_login_failed = True
                    elif retry_is_tech_error:
                        retry_status = TECH_FAILURE_STATUS
                    else:
                        ok2, err2 = selenium_go_to_inbox(driver)
                        if not ok2:
                            retry_is_tech_error = True
                            retry_error_msg = err2
                            retry_status = TECH_FAILURE_STATUS
                        else:
                            status_text, retry_is_tech_error, retry_error_msg = selenium_get_status(driver)
                            retry_status = TECH_FAILURE_STATUS if retry_is_tech_error else status_text
                except WebDriverException as exc:
                    retry_is_tech_error = True
                    retry_error_msg = f"الكروم اتقفل/وقع أثناء إعادة المحاولة: {exc}"
                    retry_status = TECH_FAILURE_STATUS
                    retry_browser_crashed = True
                finally:
                    if retry_browser_crashed:
                        driver = restart_browser(driver)
                    elif retry_login_failed:
                        clear_session(driver)
                    else:
                        selenium_logout(driver)

                update_student_status(student["id"], retry_status)
                for result in processed_results:
                    if result["name"] == display_name:
                        result["status"] = str(retry_status)
                        break

                if retry_is_tech_error:
                    print(f"    ⚠️ Still a technical error: {retry_error_msg}")
                else:
                    print(f"    ✅ Status after retry: {retry_status}")

            tech_failures = [
                failure for failure in tech_failures
                if any(
                    result["name"] == failure["name"] and result["status"] == TECH_FAILURE_STATUS
                    for result in processed_results
                )
            ]

    finally:
        try:
            driver.quit()
        except Exception:
            pass

    try:
        notify_office_status_changes(office, previous_results, processed_results)
    except Exception as exc:
        print(f"Could not send status-change email: {exc}")

    log_activity(
        job.get("office_id"),
        "اكتمل تحديث حالات الطلاب عبر Worker (Selenium + Supabase)",
        job.get("file_name") or "",
        {
            "job_id": job_id,
            "students_processed": len(processed_results),
            "technical_errors": len(tech_failures),
            "worker_id": WORKER_ID,
        },
        data_source_id=job.get("data_source_id"),
    )

    if stopped_early:
        set_job_status(
            job_id,
            "failed",
            f"تم إيقاف المهمة بعد {CONSECUTIVE_TECH_FAILURE_LIMIT} أخطاء فنية متتالية",
        )
    else:
        set_job_status(job_id, "done")

    if tech_failures:
        notify_developer_tech_failures(office, job_id, tech_failures, stopped_early=stopped_early)

    print(f"=== Job {job_id} finished ({len(processed_results)}/{total} students) ===")


# ==================== Main loop ====================
def main():
    print(f"Worker (Selenium + Supabase) running... ID: {WORKER_ID}")
    print(f"Checking for new Supabase jobs every ~{JOB_POLL_INTERVAL_SECONDS} seconds.")
    print("Chrome will open visibly as soon as a new pending job arrives.")

    while True:
        try:
            job = claim_next_pending_job()
            if job:
                process_job(job)
            else:
                time.sleep(random.uniform(
                    JOB_POLL_INTERVAL_SECONDS * 0.7,
                    JOB_POLL_INTERVAL_SECONDS * 1.3,
                ))
        except KeyboardInterrupt:
            print("Worker stopped manually.")
            break
        except Exception as exc:
            print(f"❌ Unexpected error in worker: {type(exc).__name__}: {exc!r}")
            traceback.print_exc()
            time.sleep(10)


if __name__ == "__main__":
    main()
