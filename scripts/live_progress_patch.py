from pathlib import Path

p = Path('streamlit_app (7).py')
s = p.read_text(encoding='utf-8')

# 1) Legacy API fallback: show the latest student result immediately.
needle = '''    total = len(students)\n    failed = 0\n    for index, student in enumerate(students, 1):\n'''
replacement = '''    total = len(students)\n    failed = 0\n    live_result = st.empty()\n    for index, student in enumerate(students, 1):\n'''
if needle in s and 'live_result = st.empty()' not in s:
    s = s.replace(needle, replacement, 1)

needle = '''        db().table("job_progress").insert({\n            "job_id": job["id"],\n            "student_index": index,\n            "total": total,\n            "student_name": student.get("student_name") or student.get("login_identifier"),\n            "status": status,\n        }).execute()\n\n    db().table("jobs").update({\n'''
replacement = '''        student_display = student.get("student_name") or student.get("login_identifier")\n        db().table("job_progress").insert({\n            "job_id": job["id"],\n            "student_index": index,\n            "total": total,\n            "student_name": student_display,\n            "status": status,\n        }).execute()\n        live_result.success(f"✅ تم فحص: {student_display} — الحالة الحالية: {status}")\n\n    db().table("jobs").update({\n'''
if needle in s and 'live_result.success(f"✅ تم فحص:' not in s:
    s = s.replace(needle, replacement, 1)

# 2) Worker path: highlight the latest processed student on every rerun.
needle = '''            if rows:\n                total=int(rows[-1].get("total") or 0); current=len(rows)\n                st.progress(min(current/max(total,1),1.0)); st.caption(f"طالب {current} من {total}")\n                table_df=pd.DataFrame([{"اسم الطالب":r.get("student_name","") ,"الحالة":r.get("status","")} for r in reversed(rows)])\n'''
# tolerate the exact compact table formatting currently in the file via a second simpler replacement below.
if needle not in s:
    needle = '''            if rows:\n                total=int(rows[-1].get("total") or 0); current=len(rows)\n                st.progress(min(current/max(total,1),1.0)); st.caption(f"طالب {current} من {total}")\n                table_df=pd.DataFrame([{"اسم الطالب":r.get("student_name",""),"الحالة":r.get("status","")} for r in reversed(rows)])\n'''
replacement = '''            if rows:\n                total=int(rows[-1].get("total") or 0); current=len(rows)\n                st.progress(min(current/max(total,1),1.0)); st.caption(f"طالب {current} من {total}")\n                latest = rows[-1]\n                latest_name = latest.get("student_name") or "طالب"\n                latest_status = latest.get("status") or ""\n                st.success(f"✅ آخر طالب تم فحصه: **{latest_name}** — الحالة الحالية: **{latest_status}**")\n                table_df=pd.DataFrame([{"اسم الطالب":r.get("student_name",""),"الحالة":r.get("status","")} for r in reversed(rows)])\n'''
if needle in s and 'آخر طالب تم فحصه' not in s:
    s = s.replace(needle, replacement, 1)

p.write_text(s, encoding='utf-8')
print('live progress patch applied')
