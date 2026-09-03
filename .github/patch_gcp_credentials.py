from pathlib import Path
import re

path = Path('streamlit_app (7).py')
text = path.read_text(encoding='utf-8')

new_function = '''def get_google_credentials(scopes):

    if Credentials is None:
        raise RuntimeError(
            "Google credentials libraries unavailable"
        )

    # Prefer direct Streamlit secret variables when they are available.
    # This keeps GCP credentials independent from the nested TOML section.
    direct = {
        "type": "service_account",
        "project_id": st.secrets.get(
            "GCP_PROJECT_ID",
            os.getenv("GCP_PROJECT_ID")
        ),
        "private_key_id": st.secrets.get(
            "GCP_PRIVATE_KEY_ID",
            os.getenv("GCP_PRIVATE_KEY_ID")
        ),
        "private_key": st.secrets.get(
            "GCP_PRIVATE_KEY",
            os.getenv("GCP_PRIVATE_KEY")
        ),
        "client_email": st.secrets.get(
            "GCP_CLIENT_EMAIL",
            os.getenv("GCP_CLIENT_EMAIL")
        ),
        "client_id": st.secrets.get(
            "GCP_CLIENT_ID",
            os.getenv("GCP_CLIENT_ID")
        ),
        "auth_uri": st.secrets.get(
            "GCP_AUTH_URI",
            os.getenv("GCP_AUTH_URI", "https://accounts.google.com/o/oauth2/auth")
        ),
        "token_uri": st.secrets.get(
            "GCP_TOKEN_URI",
            os.getenv("GCP_TOKEN_URI", "https://oauth2.googleapis.com/token")
        ),
    }

    if direct.get("client_email") and direct.get("private_key"):
        if isinstance(direct["private_key"], str):
            direct["private_key"] = direct["private_key"].replace("\\\\n", "\\n")
        return Credentials.from_service_account_info(
            direct,
            scopes=scopes
        )

    # Keep the existing nested secret as a fallback so Google Sheets keeps working.
    creds_dict = st.secrets.get("gcp_service_account")

    if creds_dict:
        try:
            creds_dict = dict(creds_dict)
        except Exception:
            pass

        if (
            hasattr(creds_dict, "get")
            and creds_dict.get("client_email")
            and creds_dict.get("private_key")
        ):
            return Credentials.from_service_account_info(
                creds_dict,
                scopes=scopes
            )

    raw = st.secrets.get(
        "GCP_SERVICE_ACCOUNT_JSON",
        os.getenv("GCP_SERVICE_ACCOUNT_JSON")
    )

    if raw:
        data = (
            json.loads(raw)
            if isinstance(raw, str)
            else dict(raw)
        )
        return Credentials.from_service_account_info(
            data,
            scopes=scopes
        )

    raise RuntimeError(
        "Google service account configuration missing"
    )
'''

pattern = re.compile(
r"def get_google_credentials\(scopes\):.*?(?=\ndef drive_service\(\):)",
re.DOTALL)
matches = list(pattern.finditer(text))
if not matches:
    raise SystemExit('get_google_credentials function not found')

# Replace only the active (last) definition; do not touch the earlier duplicate.
match = matches[-1]
text = text[:match.start()] + new_function + text[match.end():]
path.write_text(text, encoding='utf-8')
print('Patched active get_google_credentials:', path)
