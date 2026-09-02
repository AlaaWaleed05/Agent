from pathlib import Path

p=Path('streamlit_app (7).py')
s=p.read_text(encoding='utf-8')

hero='''_hour = datetime.now().hour
_greeting = "صباح الخير" if _hour < 12 else "مساء الخير"
st.markdown(f"""
<div class="hero">
    <div class="hero-kicker">{_greeting} 👋</div>
    <div class="hero-title">أهلاً بيك، <strong>{office.get('name', '')}</strong></div>
    <div class="hero-desc">تابع طلبات طلابك وحدّث الحالات من مكان واحد.</div>
</div>
""", unsafe_allow_html=True)
'''

top='''st.markdown("""
<div class="topbar">
    <div class="brand">
        <div class="brand-icon">✨</div>
        <div><div class="brand-title">Aivora</div><div class="brand-sub">Your Smarter Support for Every Student's Application</div></div>
    </div>
    <div style="font-size:13px;color:#6b7280;">نظام متابعة المكاتب</div>
</div>
""", unsafe_allow_html=True)
'''
if s.find(hero) != -1 and s.find(top) != -1 and s.find(hero) < s.find(top):
    s=s.replace(hero+"\n"+top, top+"\n"+hero, 1)

needle='''    if st.button("تسجيل الخروج من الإدارة", key="admin_logout"):
        reset_session_on_logout()
'''
if 'كل الحسابات' not in s and needle in s:
    block='''    st.markdown("<div class='section-title' style='margin-top:28px;'>كل الحسابات</div>", unsafe_allow_html=True)
    try:
        all_accounts = db().table("offices").select("name,email,status,created_at").order("created_at", desc=True).execute().data or []
        if all_accounts:
            import pandas as pd
            st.dataframe(pd.DataFrame(all_accounts), use_container_width=True, hide_index=True)
    except Exception:
        pass
'''
    s=s.replace(needle, block+needle, 1)

p.write_text(s,encoding='utf-8')
print('final legacy UI polish applied')
