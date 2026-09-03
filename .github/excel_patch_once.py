from pathlib import Path

APP = Path('streamlit_app (7).py')
WORKER = Path('worker.py')
WORKFLOW = Path('.github/workflows/apply-live-table.yml')
TRIGGER = Path('.github/excel_live_marker.txt')
SCRIPT = Path('.github/excel_patch_once.py')

ORIGINAL_WORKFLOW = '''name: Validate Aivora

on:
  push:
    branches: [main]
  pull_request:
    branches: [main]

permissions:
  contents: read

jobs:
  validate:
    runs-on: ubuntu-latest
    steps:
      - name: Checkout
        uses: actions/checkout@v4

      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: Compile Streamlit app
        run: python -m py_compile 'streamlit_app (7).py'

      - name: Compile Worker
        run: python -m py_compile worker.py
'''


def replace_block(text, start, end, replacement):
    a = text.find(start)
    if a < 0:
        raise RuntimeError(f'start marker not found: {start}')
    b = text.find(end, a + len(start))
    if b < 0:
        raise RuntimeError(f'end marker not found: {end}')
    return text[:a] + replacement.rstrip() + '\n\n' + text[b:]

app = APP.read_text(encoding='utf-8')
app_live = '''# =========================================================
# LIVE EXCEL UPDATE
# SAME DRIVE FILE
# =========================================================

def update_excel_student_status(source_ref, student, status):
    source_ref = _excel_text(source_ref)
    if not source_ref:
        raise RuntimeError('excel_source_missing')
    file_bytes = download_drive_file_bytes(source_ref)
    if not file_bytes:
        raise RuntimeError('excel_download_empty')

    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
    ws = wb.active
    cols, header_row = _find_excel_layout(ws)
    email_col = cols['email']
    status_col = cols.get('status')
    if status_col is None:
        status_col = ws.max_column + 1
        ws.cell(header_row, status_col).value = 'حالة الطلب'

    login = _excel_text(student.get('login_identifier')).casefold()
    source_row = student.get('source_row_number')
    target_row = None

    if source_row:
        candidate = int(source_row)
        if header_row < candidate <= ws.max_row:
            target_row = candidate
    if target_row is not None and login:
        if _excel_text(ws.cell(target_row, email_col).value).casefold() != login:
            target_row = None
    if target_row is None and login:
        for row_idx in range(header_row + 1, ws.max_row + 1):
            if _excel_text(ws.cell(row_idx, email_col).value).casefold() == login:
                target_row = row_idx
                break
    if target_row is None:
        raise RuntimeError(f'excel_student_row_not_found:{login}')

    ws.cell(target_row, status_col).value = _excel_text(status)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    updated_bytes = output.getvalue()
    if not updated_bytes:
        raise RuntimeError('excel_output_empty')

    service = drive_service()
    media = MediaIoBaseUpload(
        io.BytesIO(updated_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        resumable=False,
    )
    result = service.files().update(
        fileId=source_ref,
        media_body=media,
        supportsAllDrives=True,
        fields='id,modifiedTime,size,mimeType,name',
    ).execute()
    if str(result.get('id') or '').strip() != source_ref:
        raise RuntimeError('excel_drive_update_wrong_file')
    safe_log(f"Live Excel update: {student.get('student_name')} -> {status}")
'''
app = replace_block(app, '# =========================================================\n# LIVE EXCEL UPDATE\n# SAME DRIVE FILE\n# =========================================================', '# =========================================================\n# LIVE GOOGLE SHEET UPDATE\n# SAME SHEET\n# =========================================================', app_live)
APP.write_text(app, encoding='utf-8')

worker = WORKER.read_text(encoding='utf-8')
worker_live = '''# ============================================================
# LIVE EXCEL UPDATE
# ============================================================

def update_excel_student_status(source_ref, student, status):
    source_ref = _excel_text(source_ref)
    if not source_ref:
        raise RuntimeError('excel_source_missing')
    file_bytes = download_drive_file_bytes(source_ref)
    if not file_bytes:
        raise RuntimeError('excel_download_empty')

    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
    ws = wb.active
    cols, header_row = find_excel_columns_for_output(ws)
    status_col = find_status_column(ws, header_row)
    login = _excel_text(student.get('login_identifier')).casefold()
    source_row = student.get('source_row_number')
    target_row = None

    if source_row:
        candidate = int(source_row)
        if header_row < candidate <= ws.max_row:
            target_row = candidate
    if target_row is not None and login:
        if _excel_text(ws.cell(target_row, cols['email'] + 1).value).casefold() != login:
            target_row = None
    if target_row is None and login:
        for row_idx in range(header_row + 1, ws.max_row + 1):
            if _excel_text(ws.cell(row_idx, cols['email'] + 1).value).casefold() == login:
                target_row = row_idx
                break
    if target_row is None:
        raise RuntimeError(f'excel_student_row_not_found:{login}')

    ws.cell(target_row, status_col).value = _excel_text(status)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    updated_bytes = output.getvalue()
    if not updated_bytes:
        raise RuntimeError('excel_output_empty')

    service = drive_service()
    media = MediaIoBaseUpload(
        io.BytesIO(updated_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        resumable=False,
    )
    result = service.files().update(
        fileId=source_ref,
        media_body=media,
        supportsAllDrives=True,
        fields='id,modifiedTime,size,mimeType,name',
    ).execute()
    if str(result.get('id') or '').strip() != source_ref:
        raise RuntimeError('excel_drive_update_wrong_file')
    print(f'[EXCEL] SUCCESS file={source_ref} row={target_row} status={status}')
'''
worker = replace_block(worker, '# ============================================================\n# LIVE EXCEL UPDATE\n# ============================================================', '# ============================================================\n# LIVE GOOGLE SHEET UPDATE\n# ============================================================', worker_live)
WORKER.write_text(worker, encoding='utf-8')

# Restore the validation workflow and remove only the temporary executor files.
WORKFLOW.write_text(ORIGINAL_WORKFLOW, encoding='utf-8')
if TRIGGER.exists():
    TRIGGER.unlink()
SCRIPT.unlink()

print('Excel-only patch applied and temporary executor removed.')
