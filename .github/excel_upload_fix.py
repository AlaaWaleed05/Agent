from pathlib import Path

path = Path('streamlit_app (7).py')
text = path.read_text(encoding='utf-8')

old = '''    return (\n        service\n        .files()\n        .create(\n            body=metadata,\n            media_body=media,\n            fields="id"\n        )\n        .execute()["id"]\n    )'''

new = '''    result = (\n        service\n        .files()\n        .create(\n            body=metadata,\n            media_body=media,\n            supportsAllDrives=True,\n            fields="id"\n        )\n        .execute()\n    )\n\n    file_id = str(result.get("id") or "").strip()\n    if not file_id:\n        raise RuntimeError("excel_drive_upload_missing_id")\n\n    return file_id'''

if old not in text:
    raise SystemExit('Expected Excel upload block not found; no file changed.')

text = text.replace(old, new, 1)
path.write_text(text, encoding='utf-8')
print('Excel upload fixed: supportsAllDrives + verified file id')
