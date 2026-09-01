from pathlib import Path

p = Path('streamlit_app (7).py')
s = p.read_text(encoding='utf-8')

# Legacy API fallback: keep a live table in the same Streamlit run.
needle = '''    total = len(students)\n    failed = 0\n    for index, student in enumerate(students, 1):\n'''
replacement = '''    total = len(students)\n    failed = 0\n    live_table = st.empty()\n    live_rows = []\n    for index, student in enumerate(students, 1):\n'''
if needle in s and 'live_table = st.empty()' not in s:
    s = s.replace(needle, replacement, 1)

needle = '''        db().table("job_progress").insert({\n            "job_id": job["id"],\n            "student_index": index,\n            "total": total,\n            "student_name": student.get("student_name") or student.get("login_identifier"),\n            "status": status,\n        }).execute()\n\n    db().table("jobs").update({\n'''
replacement = '''        student_display = student.get("student_name") or student.get("login_identifier")\n        db().table("job_progress").insert({\n            "job_id": job["id"],\n            "student_index": index,\n            "total": total,\n            "student_name": student_display,\n            "status": status,\n        }).execute()\n        live_rows.append({"اسم الطالب": student_display, "الحالة الجديدة": status})\n        live_table.dataframe(\n            pd.DataFrame(live_rows),\n            use_container_width=True,\n            hide_index=True,\n        )\n\n    db().table("jobs").update({\n'''
if needle in s and '"الحالة الجديدة": status' not in s:
    s = s.replace(needle, replacement, 1)

# Worker path: keep the same table and show the most recently processed student.
needle = '''            if rows:\n                total=int(rows[-1].get("total") or 0); current=len(rows)\n                st.progress(min(current/max(total,1),1.0)); st.caption(f"طالب {current} من {total}")\n                table_df=pd.DataFrame([{"اسم الطالب":r.get("student_name",""),"الحالة":r.get("status","")} for r in reversed(rows)])\n'''
replacement = '''            if rows:\n                total=int(rows[-1].get("total") or 0); current=len(rows)\n                st.progress(min(current/max(total,1),1.0)); st.caption(f"طالب {current} من {total}")\n                latest = rows[-1]\n                latest_name = latest.get("student_name") or "طالب"\n                latest_status = latest.get("status") or ""\n                st.info(f"🔄 آخر طالب تم فحصه: **{latest_name}** — الحالة الجديدة: **{latest_status}**")\n                table_df=pd.DataFrame([{"اسم الطالب":r.get("student_name",""),"الحالة الجديدة":r.get("status","")} for r in reversed(rows)])\n'''
if needle in s and 'آخر طالب تم فحصه' not in s:
    s = s.replace(needle, replacement, 1)

p.write_text(s, encoding='utf-8')
print('live progress table patch applied')
