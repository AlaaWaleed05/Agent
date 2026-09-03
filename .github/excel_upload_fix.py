from pathlib import Path

path = Path('streamlit_app (7).py')
text = path.read_text(encoding='utf-8')

# Restore the two Google Sheets helpers that were accidentally removed.
# Insert immediately before the existing Excel helper section so that no
# unrelated application logic is replaced.
if 'def get_saved_gsheet_link(' not in text or 'def save_gsheet_link(' not in text:
    anchor = 'def _excel_text(value):'
    if anchor not in text:
        raise SystemExit('Excel helper anchor not found; no file changed.')

    block = '''def get_saved_gsheet_link(office_id):

    try:

        rows = (
            db()
            .table("data_sources")
            .select("source_url,created_at")
            .eq("office_id", office_id)
            .eq("source_type", "google_sheet")
            .not_.is_("source_url", "null")
            .order("created_at", desc=True)
            .limit(1)
            .execute()
            .data
            or []
        )

        return (
            rows[0].get("source_url")
            if rows
            else None
        )

    except Exception:
        return None


def save_gsheet_link(office_id, link):

    try:

        if not extract_sheet_id(link):
            return False, "الرابط غير صحيح!"

        (
            db()
            .table("data_sources")
            .insert({
                "office_id": office_id,
                "source_type": "google_sheet",
                "source_name": "Google Sheet",
                "source_url": link,
                "column_mapping": {},
            })
            .execute()
        )

        return True, "تم حفظ الرابط بنجاح"

    except Exception:
        safe_log("Google Sheet link save failed")
        return False, "تعذر حفظ الرابط حاليًا."


'''

    text = text.replace(anchor, block + anchor, 1)

compile(text, 'streamlit_app (7).py', 'exec')
path.write_text(text, encoding='utf-8')
print('Restored missing Google Sheets helper functions.')
# Targeted fix only; no unrelated code changes.
