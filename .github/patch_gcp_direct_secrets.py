from pathlib import Path
import ast

path = Path('streamlit_app (7).py')
text = path.read_text(encoding='utf-8')


def replace_function(source, function_name, replacement):
    tree = ast.parse(source)
    nodes = [n for n in tree.body if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef)) and n.name == function_name]
    if len(nodes) != 1:
        raise SystemExit(f'{function_name}: expected one top-level function, found {len(nodes)}')
    node = nodes[0]
    lines = source.splitlines(keepends=True)
    start = node.lineno - 1
    end = node.end_lineno
    new_lines = replacement.rstrip() + '\n\n'
    return ''.join(lines[:start]) + new_lines + ''.join(lines[end:])


normalize = '''def _excel_normalize_header(value):
    value = _excel_text(value).lower()
    value = value.replace("إ", "ا").replace("أ", "ا").replace("آ", "ا")
    value = re.sub(r"[ًٌٍَُِّْـ]", "", value)
    return re.sub(r"[^a-z0-9\u0600-\u06ff]+", "", value)
'''

kind = '''def _excel_header_kind(value):
    low = _excel_normalize_header(value)

    name_aliases = {
        "name", "fullname", "studentname", "student", "اسم", "اسماء",
        "اسمالطالب", "اسمالطالب", "الاسم"
    }
    email_aliases = {
        "email", "emailaddress", "mail", "e mail", "البريد", "البريدالالكتروني",
        "البريدااللكتروني", "البريدالإلكتروني", "الايميل", "الإيميل", "ايميل", "إيميل"
    }
    password_aliases = {
        "password", "pass", "passcode", "studentpassword", "الباسورد", "باسورد",
        "كلمةالمرور", "كلمهالمرور", "كلمةالسر", "كلمهالسر", "الرقم السري", "الرقمالسري"
    }
    status_aliases = {
        "status", "applicationstatus", "حالة", "الحالة", "حالةالطلب", "حالةالطلبات",
        "حالةالطلبة", "حالةالطلبةالجديدة", "حالةالطلبةالجديده", "حالةالطلاب",
        "حالةالطلابالجديدة", "حالةالطلابالجديده"
    }

    if low in name_aliases or "اسم" in low and "خدمة" not in low:
        return "name"
    if low in email_aliases or "email" in low or "mail" in low or "بريد" in low or "ايميل" in low:
        return "email"
    if low in password_aliases or "password" in low or "pass" in low or "باسورد" in low or "كلمةالمرور" in low or "كلمهالمرور" in low or "كلمةالسر" in low or "كلمهالسر" in low:
        return "password"
    if low in status_aliases or ("حالة" in low and "اسم" not in low and "خدمة" not in low):
        return "status"
    return None
'''

layout = '''def _find_excel_layout(ws):
    best = None

    for row_idx in range(1, min(30, ws.max_row) + 1):
        found = {}
        for col_idx, cell in enumerate(ws[row_idx], start=1):
            kind = _excel_header_kind(cell.value)
            if kind and kind not in found:
                found[kind] = col_idx

        score = (
            (3 if "email" in found else 0)
            + (3 if "password" in found else 0)
            + (1 if "name" in found else 0)
            + (1 if "status" in found else 0)
        )

        if "email" in found and "password" in found:
            return found, row_idx

        if best is None or score > best[0]:
            best = (score, found, row_idx)

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
print('Patched only Excel column detection/status-column helpers')
