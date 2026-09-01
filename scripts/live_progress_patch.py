from pathlib import Path
p = Path('streamlit_app (7).py')
s = p.read_text(encoding='utf-8')

old_css = '''[data-testid="stFileUploaderDropzone"]{background:#f8fafc!important;border:1.5px dashed #cbd5e1!important;border-radius:14px!important}[data-testid="stFileUploaderDropzone"] button{background:#fff!important;border:1px solid #bfdbfe!important;border-radius:8px!important}'''
new_css = '''[data-testid="stFileUploaderDropzone"]{background:#f8fafc!important;border:1.5px dashed #cbd5e1!important;border-radius:14px!important}[data-testid="stFileUploaderDropzone"] button{background:#fff!important;border:1px solid #bfdbfe!important;border-radius:8px!important;font-size:0!important;line-height:1!important;white-space:nowrap!important;overflow:hidden!important}[data-testid="stFileUploaderDropzone"] button span{font-size:0!important}[data-testid="stFileUploaderDropzone"] button::after{content:"اختيار ملف";font-family:'Cairo',sans-serif!important;font-size:13px!important;font-weight:700!important;line-height:1.2!important;color:#2563eb!important}'''
if old_css in s:
    s = s.replace(old_css, new_css, 1)
else:
    print('uploader CSS already patched or selector not found')

start = s.find('def get_students(office_id, search=""):\n')
end = s.find('\ndef status_class(status):', start)
if start == -1 or end == -1:
    raise SystemExit('get_students block not found')

new_func = '''def get_students(office_id, search=""):\n    rows=(\n        db().table("student_records")\n        .select("id,student_name,login_identifier,application_status,status_updated_at,source_row_number,created_at,updated_at")\n        .eq("office_id",office_id)\n        .execute().data or []\n    )\n\n    def _ts(value):\n        text=str(value or "").strip()\n        if not text:\n            return datetime.min.replace(tzinfo=timezone.utc)\n        try:\n            return datetime.fromisoformat(text.replace("Z","+00:00"))\n        except Exception:\n            return datetime.min.replace(tzinfo=timezone.utc)\n\n    latest={}\n    for row in rows:\n        key=str(row.get("login_identifier") or row.get("student_name") or "").strip().lower()\n        if not key:\n            continue\n        score=(_ts(row.get("status_updated_at")),_ts(row.get("updated_at")),_ts(row.get("created_at")))\n        current=latest.get(key)\n        if current is None or score > current[0]:\n            latest[key]=(score,row)\n    rows=[item[1] for item in latest.values()]\n    rows=sorted(rows,key=lambda r:str(r.get("student_name") or "").lower())\n    if search.strip():\n        q=search.strip().lower(); rows=[r for r in rows if q in str(r.get("student_name","")).lower()]\n    return rows\n'''

s = s[:start] + new_func + s[end:]
p.write_text(s, encoding='utf-8')
print('search and file uploader styling patched')