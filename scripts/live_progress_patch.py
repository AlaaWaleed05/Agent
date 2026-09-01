from pathlib import Path
p=Path('streamlit_app (7).py')
s=p.read_text(encoding='utf-8')
old='''def get_students(office_id, search=""):\n    rows=db().table("student_records").select("id,student_name,login_identifier,application_status,status_updated_at,source_row_number").eq("office_id",office_id).order("student_name").execute().data or []\n    if search.strip():\n        q=search.strip().lower(); rows=[r for r in rows if q in str(r.get("student_name","")).lower()]\n    return rows\n'''
new='''def get_students(office_id, search=""):\n    rows=(db().table("student_records").select("id,student_name,login_identifier,application_status,status_updated_at,source_row_number,created_at,updated_at").eq("office_id",office_id).order("updated_at",desc=True).execute().data or [])\n    latest={}\n    for row in rows:\n        key=str(row.get("login_identifier") or row.get("student_name") or "").strip().lower()\n        if key and key not in latest:\n            latest[key]=row\n    rows=sorted(latest.values(),key=lambda r:str(r.get("student_name") or "").lower())\n    if search.strip():\n        q=search.strip().lower(); rows=[r for r in rows if q in str(r.get("student_name","")).lower()]\n    return rows\n'''
if old in s:
    s=s.replace(old,new,1)
p.write_text(s,encoding='utf-8')
print('search latest-record patch refreshed')
