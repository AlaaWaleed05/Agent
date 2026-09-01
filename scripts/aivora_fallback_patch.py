from pathlib import Path

p = Path('streamlit_app (7).py')
s = p.read_text(encoding='utf-8')

marker = '# ==================== LEGACY API FALLBACK 30S ===================='
if marker not in s:
    s = s.replace('import io\nimport os\nimport re\nimport time', 'import io\nimport json\nimport os\nimport re\nimport time\n\nimport requests', 1)
    fallback = '''

# ==================== LEGACY API FALLBACK 30S ====================
BASE_URL = "https://apiadm.study-in-egypt.gov.eg/api"
SITE_URL = "https://admission.study-in-egypt.gov.eg"
WORKER_WAIT_SECONDS = 30


def _decrypt_student_password(value):
    key = st.secrets.get("STUDENT_PASSWORD_ENCRYPTION_KEY", os.getenv("STUDENT_PASSWORD_ENCRYPTION_KEY"))
    if not key:
        raise RuntimeError("STUDENT_PASSWORD_ENCRYPTION_KEY مش موجود في Secrets.")
    return Fernet(key.encode()).decrypt(str(value).encode()).decode()


def _legacy_api_login(email, password):
    session = requests.Session()
    session.headers.update({
        "accept": "application/json, text/plain, */*",
        "accept-language": "ar",
        "device": "CITIZEN",
        "origin": SITE_URL,
        "referer": SITE_URL + "/",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36",
        "content-type": "application/json",
    })
    response = session.post(f"{BASE_URL}/student/login", json={"email": email, "password": password}, timeout=30)
    if response.status_code not in (200, 201):
        return None, None, f"فشل تسجيل الدخول - كود: {response.status_code}"
    body = response.json() if response.content else {}
    token = body.get("token", "") or response.headers.get("x-csrf-token", "")
    return session, token, None


def _legacy_api_get_status(session, token):
    filt = {
        "where": {},
        "limit": 10,
        "offset": 0,
        "order": "statusUpdatedAt DESC",
        "fields": ["serviceSlug", "ID", "createdAt", "statusUpdatedAt", "activityId", "activityName"],
    }
    headers = {"x-csrf-token": token} if token else {}
    response = session.get(
        f"{BASE_URL}/dynamic_services/inbox",
        params={"filter": json.dumps(filt)},
        headers=headers,
        timeout=30,
    )
    if response.status_code not in (200, 304):
        return f"خطأ في جلب الحالة ({response.status_code})"
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
        session.post(f"{BASE_URL}/student/logout", json={"redirectUrl": SITE_URL}, timeout=15)
    except Exception:
        pass


def _claim_fallback_job(job_id):
    rows = (
        db().table("jobs")
        .update({"status": "processing", "started_at": now_iso(), "claimed_by": "streamlit-fallback"})
        .eq("id", job_id)
        .eq("status", "pending")
        .select("*")
        .execute().data or []
    )
    return rows[0] if rows else None


def _run_legacy_api_fallback(job):
    students = (
        db().table("student_records")
        .select("*")
        .eq("office_id", job["office_id"])
        .eq("data_source_id", job["data_source_id"])
        .order("source_row_number")
        .execute().data or []
    )
    students = [
        row for row in students
        if str(row.get("application_status") or "").strip() not in FINAL_STATUSES
        and row.get("login_identifier")
        and row.get("encrypted_password")
    ]
    if not students:
        db().table("jobs").update({"status": "done", "finished_at": now_iso()}).eq("id", job["id"]).execute()
        return

    total = len(students)
    failed = 0
    for index, student in enumerate(students, 1):
        session = None
        status = "خطأ فني في الفحص"
        try:
            session, token, error = _legacy_api_login(
                student["login_identifier"],
                _decrypt_student_password(student["encrypted_password"]),
            )
            if error:
                status = "فشل تسجيل الدخول"
                failed += 1
            else:
                status = _legacy_api_get_status(session, token)
                if str(status).startswith("خطأ"):
                    failed += 1
        except Exception:
            failed += 1
        finally:
            _legacy_api_logout(session)

        stamp = now_iso()
        db().table("student_records").update({
            "application_status": status,
            "status_updated_at": stamp,
            "updated_at": stamp,
        }).eq("id", student["id"]).execute()
        db().table("job_progress").insert({
            "job_id": job["id"],
            "student_index": index,
            "total": total,
            "student_name": student.get("student_name") or student.get("login_identifier"),
            "status": status,
        }).execute()

    db().table("jobs").update({
        "status": "failed" if failed >= total else "done",
        "finished_at": now_iso(),
        "error": "فشل فني في كل الطلاب" if failed >= total else None,
    }).eq("id", job["id"]).execute()


def wait_for_worker_or_legacy_fallback(job_id):
    deadline = time.monotonic() + WORKER_WAIT_SECONDS
    box = st.empty()
    progress = st.progress(0.0)
    while time.monotonic() < deadline:
        job = get_job(job_id)
        if not job:
            progress.empty()
            box.error("المهمة اختفت من Supabase.")
            return "missing"
        if str(job.get("status") or "pending") != "pending":
            progress.empty()
            return "worker"
        elapsed = int(WORKER_WAIT_SECONDS - max(0, deadline - time.monotonic()))
        remaining = max(0, WORKER_WAIT_SECONDS - elapsed)
        box.info(f"بندور على الـ Worker... لو مش موجود، التحديث هيبدأ تلقائيًا بعد {remaining} ثانية.")
        progress.progress(min(elapsed / WORKER_WAIT_SECONDS, 1.0))
        time.sleep(2)

    claimed = _claim_fallback_job(job_id)
    progress.empty()
    if claimed:
        box.warning("الـ Worker ما استلمش المهمة خلال 30 ثانية — Streamlit بدأ التحديث بالطريقة القديمة.")
        _run_legacy_api_fallback(claimed)
        return "fallback"
    return "worker"
'''
    s = s.replace('# ==================== UI ====================', fallback + '\n# ==================== UI ====================', 1)

old = '''    elif st.button("▶ تحديث حالات الطلاب",key="start_main"):
        st.session_state.update_locked=True
        try:
            is_gsheet_source=bool(saved_link and source=="🔗 ربط Google Sheets")
            source_type="gsheet" if is_gsheet_source else "xlsx"
            source_name="Google Sheet" if is_gsheet_source else (filename or "students.xlsx")
            src,count=import_students(office_id,source_type,source_name,file_bytes=file_bytes,source_url=saved_link if is_gsheet_source else None)
            job=create_job(office_id,src,filename or source_name); st.session_state.active_job_id=job["id"]
            log_activity(office_id,"إنشاء مهمة تحديث حالات",filename or source_name,{"job_id":job["id"],"students":count},data_source_id=src["id"])
            st.success(f"تم تجهيز {count} طالب. الـ Worker جاهز يستلم المهمة.")
        except Exception as exc: st.session_state.update_locked=False; st.error(str(exc))
'''
new = '''    elif st.button("▶ تحديث حالات الطلاب",key="start_main"):
        st.session_state.update_locked=True
        try:
            is_gsheet_source=bool(saved_link and source=="🔗 ربط Google Sheets")
            source_type="gsheet" if is_gsheet_source else "xlsx"
            source_name="Google Sheet" if is_gsheet_source else (filename or "students.xlsx")
            src,count=import_students(office_id,source_type,source_name,file_bytes=file_bytes,source_url=saved_link if is_gsheet_source else None)
            job=create_job(office_id,src,filename or source_name); st.session_state.active_job_id=job["id"]
            log_activity(office_id,"إنشاء مهمة تحديث حالات",filename or source_name,{"job_id":job["id"],"students":count},data_source_id=src["id"])
            st.success(f"تم تجهيز {count} طالب.")
            result=wait_for_worker_or_legacy_fallback(job["id"])
            if result=="fallback": st.success("الـ Worker ماكانش متصل، فـ Streamlit نفّذ التحديث بالطريقة القديمة.")
            elif result=="worker": st.info("الـ Worker استلم المهمة.")
        except Exception as exc:
            st.session_state.update_locked=True
            st.error(str(exc))
'''
if old not in s:
    raise SystemExit('start button block not found')
s = s.replace(old, new, 1)
p.write_text(s, encoding='utf-8')
print('fallback patch applied')
