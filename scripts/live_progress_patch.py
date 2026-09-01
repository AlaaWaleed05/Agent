from pathlib import Path

p = Path('streamlit_app (7).py')
s = p.read_text(encoding='utf-8')

# Live fallback table.
needle = '''    total = len(students)\n    failed = 0\n    for index, student in enumerate(students, 1):\n'''
replacement = '''    total = len(students)\n    failed = 0\n    live_table = st.empty()\n    live_rows = []\n    for index, student in enumerate(students, 1):\n'''
if needle in s and 'live_table = st.empty()' not in s:
    s = s.replace(needle, replacement, 1)

needle = '''        db().table("job_progress").insert({\n            "job_id": job["id"],\n            "student_index": index,\n            "total": total,\n            "student_name": student.get("student_name") or student.get("login_identifier"),\n            "status": status,\n        }).execute()\n\n    db().table("jobs").update({\n'''
replacement = '''        student_display = student.get("student_name") or student.get("login_identifier")\n        db().table("job_progress").insert({\n            "job_id": job["id"],\n            "student_index": index,\n            "total": total,\n            "student_name": student_display,\n            "status": status,\n        }).execute()\n        live_rows.append({"اسم الطالب": student_display, "الحالة الجديدة": status})\n        live_table.dataframe(pd.DataFrame(live_rows), use_container_width=True, hide_index=True)\n\n    db().table("jobs").update({\n'''
if needle in s and '"الحالة الجديدة": status' not in s:
    s = s.replace(needle, replacement, 1)

# Worker table: same columns and newest result visible.
needle = '''                table_df=pd.DataFrame([{"اسم الطالب":r.get("student_name",""),"الحالة":r.get("status","")} for r in reversed(rows)])\n                st.dataframe(table_df,use_container_width=True,hide_index=True)\n'''
replacement = '''                latest = rows[-1]\n                latest_name = latest.get("student_name") or "طالب"\n                latest_status = latest.get("status") or ""\n                st.info(f"🔄 آخر طالب تم فحصه: **{latest_name}** — الحالة الجديدة: **{latest_status}**")\n                table_df=pd.DataFrame([{"اسم الطالب":r.get("student_name",""),"الحالة الجديدة":r.get("status","")} for r in reversed(rows)])\n                st.dataframe(table_df,use_container_width=True,hide_index=True)\n'''
if needle in s and 'آخر طالب تم فحصه' not in s:
    s = s.replace(needle, replacement, 1)

# Search: only the latest saved record per student/login identifier.
old = '''def get_students(office_id, search=""):\n    rows=db().table("student_records").select("id,student_name,login_identifier,application_status,status_updated_at,source_row_number").eq("office_id",office_id).order("student_name").execute().data or []\n    if search.strip():\n        q=search.strip().lower(); rows=[r for r in rows if q in str(r.get("student_name","")).lower()]\n    return rows\n'''
new = '''def get_students(office_id, search=""):\n    rows = (\n        db().table("student_records")\n        .select("id,student_name,login_identifier,application_status,status_updated_at,source_row_number,created_at,updated_at")\n        .eq("office_id", office_id)\n        .order("updated_at", desc=True)\n        .execute().data or []\n    )\n    # Each update imports a fresh source, so keep only the newest record for each student.\n    latest = {}\n    for row in rows:\n        key = str(row.get("login_identifier") or row.get("student_name") or "").strip().lower()\n        if key and key not in latest:\n            latest[key] = row\n    rows = sorted(latest.values(), key=lambda r: str(r.get("student_name") or "").lower())\n    if search.strip():\n        q = search.strip().lower()\n        rows = [r for r in rows if q in str(r.get("student_name", "")).lower()]\n    return rows\n'''
if old in s:
    s = s.replace(old, new, 1)

# Restore the office-page Aivora header above the data source card.
needle = '''office_id=office["id"]\n\n# ==================== Data source card ====================\n'''
replacement = '''office_id=office["id"]\n\nst.markdown("<div class='topbar'><div class='brand'><div class='brand-icon'>✨</div><div><div class='brand-title'>Aivora</div><div class='brand-sub'>Your Smarter Support for Every Student's Application</div></div></div></div>", unsafe_allow_html=True)\n\n# ==================== Data source card ====================\n'''
if needle in s and '<div class=\'topbar\'><div class=\'brand\'>' not in s:
    s = s.replace(needle, replacement, 1)

# Excel option intentionally has exactly one uploader. Keep the Google-Sheet replacement uploader only inside its own branch.
needle = '''if source=="📂 رفع ملف Excel":\n    uploaded=st.file_uploader("ارفع ملف Excel",type=["xlsx","xls"],label_visibility="collapsed",key="excel_upload")\n    if uploaded: file_bytes=uploaded.getvalue(); filename=uploaded.name; st.success(f"تم اختيار الملف: {uploaded.name}")\n'''
replacement = '''if source=="📂 رفع ملف Excel":\n    uploaded = st.file_uploader(\n        "ارفع ملف Excel",\n        type=["xlsx", "xls"],\n        label_visibility="collapsed",\n        key="excel_upload",\n    )\n    if uploaded:\n        file_bytes = uploaded.getvalue()\n        filename = uploaded.name\n        st.success(f"تم اختيار الملف: {uploaded.name}")\n'''
if needle in s:
    s = s.replace(needle, replacement, 1)

p.write_text(s, encoding='utf-8')
print('Aivora UI/search/live-table patch applied')
