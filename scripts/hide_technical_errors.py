from pathlib import Path

APP = Path("streamlit_app (7).py")
text = APP.read_text(encoding="utf-8")

# Keep user-facing validation errors, but never expose raw technical details.
# Streamlit itself is also configured separately at deployment level; this
# wrapper protects explicit st.error()/st.exception() calls in the app.
if "def safe_st_error(message):" not in text:
    marker = "ADMIN_USERNAME = st.secrets.get(\"ADMIN_USERNAME\", \"admin\")\n"
    helper = '''\n\ndef safe_st_error(message):\n    \"\"\"Show a safe message to the office and keep technical details internal.\"\"\"\n    text = str(message or \"\")\n    lower = text.lower()\n    technical_markers = (\n        \"traceback\", \"exception\", \"keyerror\", \"typeerror\", \"attributeerror\",\n        \"webdriver\", \"selenium\", \"supabase\", \"http\", \"status code\",\n        \"connection\", \"timeout\", \"noneType\", \"jsondecode\", \"request\",\n        \"postgrest\", \"psycopg\", \"ssl\", \"dns\", \"forbidden\", \"unauthorized\",\n        \"internal server error\", \"500\", \"502\", \"503\", \"504\",\n    )\n    looks_technical = any(marker in lower for marker in technical_markers)\n    if looks_technical:\n        error_id = f\"ERR-{__import__('uuid').uuid4().hex[:8].upper()}\"\n        print(f\"[AIVORA {error_id}] {text}\")\n        st.error(f\"حصلت مشكلة مؤقتة أثناء تنفيذ العملية. رقم المتابعة: {error_id}\")\n        return\n    st.error(message)\n\n'''
    if marker not in text:
        raise SystemExit("Could not find insertion marker")
    text = text.replace(marker, marker + helper, 1)

# Redirect all explicit Streamlit error renderers through the safe wrapper.
text = text.replace("st.exception(", "safe_st_error(")
# Avoid replacing the implementation call we just inserted.
parts = text.split("def safe_st_error(message):", 1)
if len(parts) == 2:
    head, tail = parts
    tail = tail.replace("st.error(", "__SAFE_ST_ERROR__(", 1)
    text = head + "def safe_st_error(message):" + tail
text = text.replace("st.error(", "safe_st_error(")
text = text.replace("__SAFE_ST_ERROR__(", "st.error(")

APP.write_text(text, encoding="utf-8")
print("technical error shielding applied")
