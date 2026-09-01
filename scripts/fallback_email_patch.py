from pathlib import Path

APP = Path("streamlit_app (7).py")
text = APP.read_text(encoding="utf-8")

old = '''    total = len(students)\n    live_table = st.empty()\n    live_rows = []\n    technical_failures = []\n'''
new = '''    total = len(students)\n    live_table = st.empty()\n    live_rows = []\n    technical_failures = []\n    previous_results = {\n        str(row.get("student_name") or row.get("login_identifier") or "").strip():\n        str(row.get("application_status") or "").strip()\n        for row in students\n        if str(row.get("student_name") or row.get("login_identifier") or "").strip()\n    }\n    processed_results = []\n'''
if old not in text:
    raise SystemExit("fallback setup block not found")
text = text.replace(old, new, 1)

old = '''        live_rows.append({"اسم الطالب": student_display, "الحالة الجديدة": status})\n        live_table.dataframe(pd.DataFrame(live_rows), use_container_width=True, hide_index=True)\n'''
new = '''        live_rows.append({"اسم الطالب": student_display, "الحالة الجديدة": status})\n        processed_results.append({"name": str(student_display).strip(), "status": str(status)})\n        live_table.dataframe(pd.DataFrame(live_rows), use_container_width=True, hide_index=True)\n'''
if old not in text:
    raise SystemExit("fallback save_result block not found")
text = text.replace(old, new, 1)

old = '''    db().table("jobs").update({\n        "status": "failed" if len(remaining_tech) >= total else "done",\n        "finished_at": now_iso(),\n        "error": "فشل فني في كل الطلاب" if len(remaining_tech) >= total else None,\n    }).eq("id", job["id"]).execute()\n'''
new = '''    # نفس إشعار الـWorker: أرسل للمكتب فقط الطلاب الذين تغيّرت حالتهم.\n    try:\n        office = get_office(job.get("office_id"))\n        notify_office_status_changes(office, previous_results, processed_results)\n    except Exception as exc:\n        print(f"Could not send fallback status-change email: {exc}")\n\n    db().table("jobs").update({\n        "status": "failed" if len(remaining_tech) >= total else "done",\n        "finished_at": now_iso(),\n        "error": "فشل فني في كل الطلاب" if len(remaining_tech) >= total else None,\n    }).eq("id", job["id"]).execute()\n'''
if old not in text:
    raise SystemExit("fallback final job-update block not found")
text = text.replace(old, new, 1)

APP.write_text(text, encoding="utf-8")
print("patched")
