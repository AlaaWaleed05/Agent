from pathlib import Path
p=Path('streamlit_app (7).py')
s=p.read_text(encoding='utf-8')

start=s.find('def get_students(office_id, search=""):\n')
end=s.find('\ndef status_class(status):', start)
if start == -1 or end == -1:
    raise SystemExit('get_students block not found')

new_func='''def get_students(office_id, search=""):\n    rows=(\n        db().table("student_records")\n        .select("id,student_name,login_identifier,application_status,status_updated_at,source_row_number,created_at,updated_at")\n        .eq("office_id",office_id)\n        .execute().data or []\n    )\n\n    def _ts(value):\n        text=str(value or "").strip()\n        if not text:\n            return datetime.min.replace(tzinfo=timezone.utc)\n        try:\n            return datetime.fromisoformat(text.replace("Z","+00:00"))\n        except Exception:\n            return datetime.min.replace(tzinfo=timezone.utc)\n\n    # لكل طالب: نعرض أحدث حالة فحص فعلية أولاً.\n    # لو فيه سجل أحدث من غير حالة (مثلاً نتيجة رفع جديد)، ما يخليش البحث\n    # يرجع "لم يتم الفحص بعد" فوق حالة أحدث محفوظة بالفعل.\n    latest={}\n    for row in rows:\n        key=str(row.get("login_identifier") or row.get("student_name") or "").strip().lower()\n        if not key:\n            continue\n        score=(\n            _ts(row.get("status_updated_at")),\n            _ts(row.get("updated_at")),\n            _ts(row.get("created_at")),\n        )\n        current=latest.get(key)\n        if current is None or score > current[0]:\n            latest[key]=(score,row)\n\n    rows=[item[1] for item in latest.values()]\n    rows=sorted(rows,key=lambda r:str(r.get("student_name") or "").lower())\n    if search.strip():\n        q=search.strip().lower()\n        rows=[r for r in rows if q in str(r.get("student_name","")).lower()]\n    return rows\n'''

s=s[:start]+new_func+s[end:]
p.write_text(s,encoding='utf-8')
print('search now prefers latest actual status update')
