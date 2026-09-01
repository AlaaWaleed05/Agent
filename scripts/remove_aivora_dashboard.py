from pathlib import Path
import re

p = Path('streamlit_app (7).py')
s = p.read_text(encoding='utf-8')

pattern = r"\n# ==================== Dashboard ====================\noffice=st\.session_state\.office\n.*?st\.markdown\(\"<div style='height:18px'></div>\",unsafe_allow_html=True\)\n"
new, count = re.subn(pattern, "\n# ==================== Office page ====================\noffice=st.session_state.office\nif not office: st.session_state.logged_in=False; st.rerun()\noffice_id=office[\"id\"]\n\n", s, count=1, flags=re.S)
if count != 1:
    raise SystemExit(f'dashboard block not found: {count}')
p.write_text(new, encoding='utf-8')
print('dashboard removed')
