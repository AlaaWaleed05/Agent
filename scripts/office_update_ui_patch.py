from pathlib import Path

APP = Path("streamlit_app (7).py")
text = APP.read_text(encoding="utf-8")

old = '''                st.success(f"تم تجهيز {count} طالب وبدأ التحديث.")\n                st.rerun()\n'''
new = '''                st.info(f"تم تجهيز {count} طالب. سيبدأ التحديث خلال ثواني…")\n                st.rerun()\n'''
if old not in text:
    raise SystemExit("Could not find start-update message")
text = text.replace(old, new, 1)

old = '''        if status=="pending": st.info("⏳ بندور على الـ Worker... لو مش موجود، التحديث هيبدأ بالطريقة القديمة بعد 30 ثانية.")\n        elif status=="processing": st.info("🔄 التحديث شغال — الـ Worker أو Streamlit بيحدّث الحالات في الخلفية.")\n'''
new = '''        if status=="pending": st.info("⏳ جاري تجهيز التحديث… سيبدأ خلال ثواني.")\n        elif status=="processing": st.success("▶️ بدأ التحديث. جاري فحص الطلاب وتحديث الحالات…")\n'''
if old not in text:
    raise SystemExit("Could not find worker-facing status messages")
text = text.replace(old, new, 1)

old = '''            last=rows[-1]; st.info(f"🔄 آخر طالب تم فحصه: **{last.get('student_name') or 'طالب'}** — الحالة الجديدة: **{last.get('status') or ''}**")\n            st.dataframe(pd.DataFrame([{\"اسم الطالب\":r.get(\"student_name\",\"\"),\"الحالة الجديدة\":r.get(\"status\",\"\")} for r in reversed(shown)]),use_container_width=True,hide_index=True)\n'''
new = '''            last=shown[-1]; st.info(f"🔄 آخر طالب تم فحصه: **{last.get('student_name') or 'طالب'}** — الحالة الجديدة: **{last.get('status') or ''}**")\n\n            table_rows=[]\n            for row_number, r in enumerate(reversed(shown), 1):\n                name = str(r.get("student_name") or "").strip()\n                status_text = str(r.get("status") or "").strip() or "لم يتم الفحص بعد"\n                table_rows.append(\n                    f"<tr><td class='office-index'>{row_number}</td><td class='office-name'>{name}</td><td class='office-status'>{status_text}</td></tr>"\n                )\n            st.markdown(\n                "<div class='office-status-table-wrap'><table class='office-status-table'><thead><tr><th>#</th><th>اسم الطالب</th><th>الحالة</th></tr></thead><tbody>"\n                + "".join(table_rows)\n                + "</tbody></table></div>",\n                unsafe_allow_html=True,\n            )\n'''
if old not in text:
    raise SystemExit("Could not find progress table block")
text = text.replace(old, new, 1)

old = '''            if st.button("تسجيل الخروج",key="logout_main"): st.session_state.clear(); st.rerun()\n'''
new = '''            if st.button("تسجيل الخروج",key="logout_main"): st.session_state.clear(); st.rerun()\n'''
# no-op anchor kept intentionally

marker = '''.result-card{background:#fff;border:1px solid #e5e7eb;border-radius:13px;padding:14px 16px;margin:8px 0}.result-name'''
insert = '''.office-status-table-wrap{overflow-x:auto;border:1px solid #e5e7eb;border-radius:14px;background:#fff;margin-top:12px}.office-status-table{width:100%;border-collapse:separate;border-spacing:0;direction:rtl;font-family:'Cairo',sans-serif!important}.office-status-table th,.office-status-table td{padding:11px 14px;border-bottom:1px solid #eef2f7;text-align:right;vertical-align:middle;font-family:'Cairo',sans-serif!important}.office-status-table thead th{background:#f8fafc;color:#374151;font-size:13px;font-weight:800}.office-status-table tbody td{color:#111827;font-size:13px}.office-status-table tbody tr:last-child td{border-bottom:none}.office-status-table .office-index{width:58px;text-align:center!important;color:#6b7280;font-weight:700}.office-status-table .office-name{font-weight:800;min-width:240px}.office-status-table .office-status{color:#2563eb;font-weight:700;min-width:220px}'''
if marker not in text:
    raise SystemExit("Could not find CSS insertion point")
text = text.replace(marker, insert + marker, 1)

APP.write_text(text, encoding="utf-8")
print("office update UI patch applied")
