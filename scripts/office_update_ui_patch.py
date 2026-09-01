from pathlib import Path
import re

APP = Path("streamlit_app (7).py")
text = APP.read_text(encoding="utf-8")

# Office-facing messages: never mention internal Worker/fallback implementation.
text, n = re.subn(
    r'st\.(?:success|info)\(f?"تم تجهيز \{count\} طالب(?: وبدأ التحديث)?\."\)',
    'st.info(f"تم تجهيز {count} طالب. سيبدأ التحديث خلال ثواني…")',
    text,
    count=1,
)
if n == 0:
    raise SystemExit("Could not find start-update message")

text = text.replace(
    'if status=="pending": st.info("⏳ بندور على الـ Worker... لو مش موجود، التحديث هيبدأ بالطريقة القديمة بعد 30 ثانية.")',
    'if status=="pending": st.info("⏳ جاري تجهيز التحديث… سيبدأ خلال ثواني.")',
    1,
)
text = text.replace(
    'elif status=="processing": st.info("🔄 التحديث شغال — الـ Worker أو Streamlit بيحدّث الحالات في الخلفية.")',
    'elif status=="processing": st.success("▶️ بدأ التحديث. جاري فحص الطلاب وتحديث الحالات…")',
    1,
)

# Replace the progress dataframe with an RTL numbered table.
marker_patterns = [
    r'(?ms)^\s*last=(?:rows|shown)\[-1\]; st\.info\(f"🔄 آخر طالب تم فحصه:.*?^\s*st\.dataframe\(pd\.DataFrame\(\[\{.*?\n',
    r'(?ms)^\s*latest=rows\[-1\]; st\.info\(f"🔄 آخر طالب تم فحصه:.*?^\s*st\.dataframe\(pd\.DataFrame\(\[\{.*?\n',
]

progress_replacement = '''            last=shown[-1] if shown else rows[-1]\n            st.info(f"🔄 آخر طالب تم فحصه: **{last.get('student_name') or 'طالب'}** — الحالة الجديدة: **{last.get('status') or 'لم يتم الفحص بعد'}**")\n\n            table_rows=[]\n            for row_number, r in enumerate(reversed(shown), 1):\n                name = str(r.get("student_name") or "").strip()\n                status_text = str(r.get("status") or "").strip() or "لم يتم الفحص بعد"\n                table_rows.append(\n                    f"<tr><td class='office-index'>{row_number}</td><td class='office-name'>{name}</td><td class='office-status'>{status_text}</td></tr>"\n                )\n            st.markdown(\n                "<div class='office-status-table-wrap'><table class='office-status-table'><thead><tr><th>#</th><th>اسم الطالب</th><th>الحالة</th></tr></thead><tbody>"\n                + "".join(table_rows)\n                + "</tbody></table></div>",\n                unsafe_allow_html=True,\n            )\n'''

replaced = False
for pattern in marker_patterns:
    text2, n = re.subn(pattern, progress_replacement, text, count=1)
    if n:
        text = text2
        replaced = True
        break
if not replaced:
    raise SystemExit("Could not find progress table block")

# Insert table CSS once.
css_marker = '.result-card{background:#fff;border:1px solid #e5e7eb;border-radius:13px;padding:14px 16px;margin:8px 0}.result-name'
css_insert = ".office-status-table-wrap{overflow-x:auto;border:1px solid #e5e7eb;border-radius:14px;background:#fff;margin-top:12px}.office-status-table{width:100%;border-collapse:separate;border-spacing:0;direction:rtl;font-family:'Cairo',sans-serif!important}.office-status-table th,.office-status-table td{padding:11px 14px;border-bottom:1px solid #eef2f7;text-align:right;vertical-align:middle;font-family:'Cairo',sans-serif!important}.office-status-table thead th{background:#f8fafc;color:#374151;font-size:13px;font-weight:800}.office-status-table tbody td{color:#111827;font-size:13px}.office-status-table tbody tr:last-child td{border-bottom:none}.office-status-table .office-index{width:58px;text-align:center!important;color:#6b7280;font-weight:700}.office-status-table .office-name{font-weight:800;min-width:240px}.office-status-table .office-status{color:#2563eb;font-weight:700;min-width:220px}"
if 'office-status-table-wrap' not in text:
    if css_marker not in text:
        raise SystemExit("Could not find CSS insertion point")
    text = text.replace(css_marker, css_insert + css_marker, 1)

APP.write_text(text, encoding="utf-8")
print("office update UI patch applied")
