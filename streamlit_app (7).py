def update_excel_student_status(
    source_ref,
    student,
    status
):

    if not source_ref:
        raise RuntimeError(
            "excel_source_missing"
        )

    file_bytes = download_excel_from_storage(
        source_ref
    )

    wb = openpyxl.load_workbook(
        io.BytesIO(file_bytes),
        data_only=False
    )
    ws = wb.active

    cols, header_row = find_excel_columns(ws)
    status_col = find_status_column_for_output(
        ws,
        header_row
    )

    source_row = student.get("source_row_number")

    if not source_row:
        login = str(
            student.get("login_identifier") or ""
        ).strip().lower()

        for row_idx in range(
            header_row + 1,
            ws.max_row + 1
        ):
            email = str(
                ws.cell(
                    row_idx,
                    cols["email"] + 1
                ).value or ""
            ).strip().lower()

            if email == login:
                source_row = row_idx
                break

    if not source_row:
        raise RuntimeError(
            "excel_student_row_missing"
        )

    ws.cell(
        int(source_row),
        status_col
    ).value = status

    output = io.BytesIO()
    wb.save(output)
    updated_bytes = output.getvalue()

    update_excel_in_storage(
        source_ref,
        updated_bytes
    )
