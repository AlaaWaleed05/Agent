from pathlib import Path
import ast

path = Path('streamlit_app (7).py')
text = path.read_text(encoding='utf-8')


def replace_function(source, function_name, replacement):
    tree = ast.parse(source)
    nodes = [
        node
        for node in ast.walk(tree)
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
        and node.name == function_name
    ]
    if len(nodes) != 1:
        raise SystemExit(
            f'{function_name}: expected one function, found {len(nodes)}'
        )
    node = nodes[0]
    lines = source.splitlines(keepends=True)
    start = node.lineno - 1
    end = node.end_lineno
    return ''.join(lines[:start]) + replacement.rstrip() + '\n\n' + ''.join(lines[end:])


normalize = '''def _excel_normalize_header(value):
    value = _excel_text(value).lower()
    value = value.replace("إ", "ا").replace("أ", "ا").replace("آ", "ا")
    value = re.sub(r"[ًٌٍَُِّْـ]", "", value)
    return re.sub(r"[^a-z0-9\u0600-\u06ff]+", "", value)
'''

kind = '''def _excel_header_kind(value):
    low = _excel_normalize_header(value)

    if low in {"name", "fullname", "studentname", "student", "اسم", "الاسم", "اسمالطالب"} or ("اسم" in low and "خدمة" not in low):
        return "name"

    if low in {"email", "emailaddress", "mail", "البريد", "البريدالالكتروني", "البريدالإلكتروني", "الايميل", "الإيميل", "ايميل", "إيميل"} or any(token in low for token in ("email", "mail", "بريد", "ايميل")):
        return "email"

    if low in {"password", "pass", "studentpassword", "الباسورد", "باسورد", "كلمةالمرور", "كلمهالمرور", "كلمةالسر", "كلمهالسر", "الرقمالسري"} or any(token in low for token in ("password", "pass", "باسورد", "كلمةالمرور", "كلمهالمرور", "كلمةالسر", "كلمهالسر")):
        return "password"

    if low in {"status", "applicationstatus", "حالة", "الحالة", "حالةالطلب", "حالةالطلبات", "حالةالطلبة", "حالةالطلبةالجديدة", "حالةالطلبةالجديده", "حالةالطلاب", "حالةالطلابالجديدة", "حالةالطلابالجديده"} or ("حالة" in low and "اسم" not in low and "خدمة" not in low):
        return "status"

    return None
'''

layout = '''def _find_excel_layout(ws):
    for row_idx in range(1, min(30, ws.max_row) + 1):
        found = {}
        for col_idx, cell in enumerate(ws[row_idx], start=1):
            kind = _excel_header_kind(cell.value)
            if kind and kind not in found:
                found[kind] = col_idx
        if "email" in found and "password" in found:
            return found, row_idx

    raise ValueError(
        "ملف Excel لازم يحتوي على أعمدة الإيميل والباسورد."
    )
'''

status = '''def find_status_column_for_output(ws, header_row):
    for col_idx, cell in enumerate(ws[header_row], start=1):
        if _excel_header_kind(cell.value) == "status":
            return col_idx

    new_col = ws.max_column + 1
    ws.cell(header_row, new_col).value = "حالة الطلبة الجديدة"
    return new_col
'''

for name, replacement in [
    ('_excel_normalize_header', normalize),
    ('_excel_header_kind', kind),
    ('_find_excel_layout', layout),
    ('find_status_column_for_output', status),
]:
    text = replace_function(text, name, replacement)

ast.parse(text)
path.write_text(text, encoding='utf-8')
print('Patched only Excel column detection/status helpers')
