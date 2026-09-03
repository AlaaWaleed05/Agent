from pathlib import Path

p = Path('worker.py')
text = p.read_text(encoding='utf-8')

if 'def update_excel_student_status(' not in text:
    marker = '# ============================================================\n# LIVE SOURCE UPDATE\n# ============================================================'
    fn = '''# ============================================================
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
        row_login = _excel_text(ws.cell(target_row, cols['email'] + 1).value).casefold()
        if row_login != login:
            target_row = None

    if target_row is None and login:
        for row_idx in range(header_row + 1, ws.max_row + 1):
            row_login = _excel_text(ws.cell(row_idx, cols['email'] + 1).value).casefold()
            if row_login == login:
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
    if marker not in text:
        raise RuntimeError('LIVE SOURCE UPDATE marker not found in worker.py')
    text = text.replace(marker, fn + marker, 1)
    p.write_text(text, encoding='utf-8')
    print('WORKER_EXCEL_FUNCTION_ADDED')
else:
    print('WORKER_EXCEL_FUNCTION_ALREADY_PRESENT')
