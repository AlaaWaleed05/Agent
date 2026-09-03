from pathlib import Path

path = Path('streamlit_app (7).py')
text = path.read_text(encoding='utf-8')

old = '''def get_google_credentials(scopes):

    if Credentials is None:
        raise RuntimeError(
            "Google credentials libraries unavailable"
        )

    creds_dict = st.secrets.get(
        "gcp_service_account"
    )

    if isinstance(creds_dict, dict):

        return (
            Credentials
            .from_service_account_info(
                creds_dict,
                scopes=scopes
            )
        )

    raw = st.secrets.get(
        "GCP_SERVICE_ACCOUNT_JSON",
        os.getenv("GCP_SERVICE_ACCOUNT_JSON")
    )

    if raw:

        return (
            Credentials
            .from_service_account_info(
                json.loads(raw),
                scopes=scopes
            )
        )

    raise RuntimeError(
        "Google service account configuration missing"
    )
'''

new = '''def get_google_credentials(scopes):

    if Credentials is None:
        raise RuntimeError(
            "Google credentials libraries unavailable"
        )

    # Prefer direct Streamlit secret variables when present.
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
            os.getenv(
                "GCP_AUTH_URI",
                "https://accounts.google.com/o/oauth2/auth"
            )
        ),
        "token_uri": st.secrets.get(
            "GCP_TOKEN_URI",
            os.getenv(
                "GCP_TOKEN_URI",
                "https://oauth2.googleapis.com/token"
            )
        ),
    }

    if direct.get("client_email") and direct.get("private_key"):
        if isinstance(direct["private_key"], str):
            direct["private_key"] = direct["private_key"].replace(
                "\\\\n",
                "\\n"
            )
        return Credentials.from_service_account_info(
            direct,
            scopes=scopes
        )

    # Existing nested secret remains a fallback for Google Sheets/Drive.
    creds_dict = st.secrets.get(
        "gcp_service_account"
    )

    if isinstance(creds_dict, dict):

        return (
            Credentials
            .from_service_account_info(
                creds_dict,
                scopes=scopes
            )
        )

    raw = st.secrets.get(
        "GCP_SERVICE_ACCOUNT_JSON",
        os.getenv("GCP_SERVICE_ACCOUNT_JSON")
    )

    if raw:

        return (
            Credentials
            .from_service_account_info(
                json.loads(raw),
                scopes=scopes
            )
        )

    raise RuntimeError(
        "Google service account configuration missing"
    )
'''

needle = '\n\ndef drive_service():'
end = text.find(needle)
if end < 0:
    raise SystemExit('drive_service boundary not found')
start = text.rfind('\ndef get_google_credentials(scopes):', 0, end)
if start < 0:
    raise SystemExit('active get_google_credentials not found')

current = text[start + 1:end]
if current != old:
    raise SystemExit('active get_google_credentials did not match expected original; refusing to modify file')

text = text[:start + 1] + new + text[end:]
path.write_text(text, encoding='utf-8')
print('Patched only the active get_google_credentials function')
