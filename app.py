#!/usr/bin/env python3
"""
HKPO Veterans Academy – Acceptance Letter & Training Plan Generator
Streamlit Web Version (Stage 2: Generate + Drive upload + folder move)
"""

import os
import re
import shutil
import subprocess
import tempfile
import json
from datetime import datetime, date
from pathlib import Path
from io import BytesIO

import streamlit as st

try:
    from docx import Document
except ImportError:
    st.error("Missing dependency: python-docx")
    st.stop()

# Google libraries (optional until login)
try:
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import Flow
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseUpload
    GOOGLE_AVAILABLE = True
except ImportError:
    GOOGLE_AVAILABLE = False

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
TEMPLATE_ROOT = SCRIPT_DIR / "hkpo_templates"
OUTPUT_DIR = SCRIPT_DIR / "output"

# Google Drive folder IDs
SUBMISSIONS_FOLDER_ID = "12MfLDc8Jb4m6b5ke6NaEdiN433SQ0MRV"
ACCEPTED_FOLDER_ID = "1SQ_yDzpWCdzKfl6K_QA2UQWklCjipLYz"

# OAuth scopes
SCOPES = [
    "https://www.googleapis.com/auth/drive",
]

# Available templates: (duration_days, filename)
TEMPLATES = {
    "Lean Practitioner": {
        False: [(60, "lean_practitioner/60_day.docx")],
        True: [],
    },
    "Green Belt": {
        False: [
            (60, "green_belt/60_day.docx"),
            (63, "green_belt/63_day.docx"),
            (88, "green_belt/88_day.docx"),
            (90, "green_belt/90_day.docx"),
        ],
        True: [
            (60, "green_belt_pmp/60_day.docx"),
            (73, "green_belt_pmp/73_day.docx"),
            (92, "green_belt_pmp/92_day.docx"),
        ],
    },
    "Black Belt": {
        False: [
            (90, "black_belt/90_day.docx"),
            (120, "black_belt/120_day.docx"),
        ],
        True: [
            (89, "black_belt_pmp/89_day.docx"),
            (90, "black_belt_pmp/90_day.docx"),
            (105, "black_belt_pmp/105_day.docx"),
            (120, "black_belt_pmp/120_day.docx"),
            (180, "black_belt_pmp/180_day.docx"),
        ],
    },
}

ACCEPTANCE_TEMPLATE = TEMPLATE_ROOT / "acceptance" / "acceptance_letter.docx"


# ---------------------------------------------------------------------------
# Document generation helpers (unchanged logic)
# ---------------------------------------------------------------------------
def find_closest_template(belt: str, pmp: bool, approx_days: int):
    options = TEMPLATES.get(belt, {}).get(pmp, [])
    if not options:
        raise ValueError(
            f"No templates available for {belt} {'+ PMP' if pmp else '(no PMP)'}."
        )
    return min(options, key=lambda x: abs(x[0] - approx_days))


def format_date_long(dt: datetime) -> str:
    return dt.strftime("%B %d, %Y").replace(" 0", " ")


def format_acceptance_dates(start: datetime, end: datetime) -> str:
    def part(d):
        return d.strftime("%B %d, %Y").replace(" 0", " ")
    return f"{part(start)} through {part(end)}"


def inclusive_days(start: datetime, end: datetime) -> int:
    return (end - start).days + 1


def replace_in_paragraph(paragraph, replacements: dict):
    if not paragraph.text:
        return
    full = paragraph.text
    new_text = full
    for old, new in replacements.items():
        if old in new_text:
            new_text = new_text.replace(old, new)
    if new_text == full:
        return
    for run in paragraph.runs:
        run.text = ""
    if paragraph.runs:
        paragraph.runs[0].text = new_text
    else:
        paragraph.add_run(new_text)


def replace_everywhere(doc: Document, replacements: dict):
    for p in doc.paragraphs:
        replace_in_paragraph(p, replacements)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p, replacements)


def generate_acceptance(first: str, last: str, start: datetime, end: datetime) -> bytes:
    if not ACCEPTANCE_TEMPLATE.exists():
        raise FileNotFoundError(f"Acceptance letter template not found: {ACCEPTANCE_TEMPLATE}")
    doc = Document(ACCEPTANCE_TEMPLATE)
    full_name = f"{first} {last}"
    date_str = format_acceptance_dates(start, end)
    replace_everywhere(doc, {
        "First Last": full_name,
        "Month XX, 20XX through Month XX, 20XX": date_str,
    })
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def generate_training_plan(belt: str, pmp: bool, approx_days: int,
                           first: str, last: str,
                           start: datetime, end: datetime,
                           co_director: str) -> bytes:
    chosen_days, rel_path = find_closest_template(belt, pmp, approx_days)
    src = TEMPLATE_ROOT / rel_path
    if not src.exists():
        raise FileNotFoundError(f"Template not found: {src}")

    doc = Document(src)
    full_name = f"{first} {last}"
    date_str = f"{format_date_long(start)} – {format_date_long(end)}"
    actual_days = inclusive_days(start, end)

    for p in doc.paragraphs:
        if p.text.strip().startswith("Dear "):
            for run in p.runs:
                run.text = ""
            if p.runs:
                p.runs[0].text = f"Dear {full_name},"
            else:
                p.add_run(f"Dear {full_name},")
            break

    for p in doc.paragraphs:
        txt = p.text
        if "Course Dates" in txt or "Course dates" in txt:
            new = re.sub(r"(Course Dates[:\s]*)(.*)", rf"\1{date_str}", txt, flags=re.IGNORECASE)
            for run in p.runs:
                run.text = ""
            if p.runs:
                p.runs[0].text = new
            else:
                p.add_run(new)

    for p in doc.paragraphs:
        txt = p.text
        if "Length of Training" in txt:
            new = re.sub(r"(Length of Training[:\s]*)(.*)", rf"\g<1>{actual_days} days", txt, flags=re.IGNORECASE)
            for run in p.runs:
                run.text = ""
            if p.runs:
                p.runs[0].text = new
            else:
                p.add_run(new)

    for p in doc.paragraphs:
        stripped = p.text.strip()
        if stripped in ("Will Roberts", "William Wingertsahn", "First Last", "Clay Gober", "Matthew Hale"):
            for run in p.runs:
                run.text = ""
            if p.runs:
                p.runs[0].text = co_director
            else:
                p.add_run(co_director)

    replace_everywhere(doc, {"First Last": full_name})

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# Google OAuth + Drive helpers
# ---------------------------------------------------------------------------
def get_redirect_uri():
    """Return the current app URL for OAuth redirect."""
    try:
        # Streamlit Cloud sets this
        return st.get_option("server.baseUrlPath") or ""
    except Exception:
        pass
    # Fallback – user should have set the correct redirect URI
    return "https://hkpo-acceptance-generator-ucqvpysmj6jpdsb4a4knny.streamlit.app/"


def create_flow():
    client_id = st.secrets["google"]["client_id"]
    client_secret = st.secrets["google"]["client_secret"]

    # Build a client config dict
    client_config = {
        "web": {
            "client_id": client_id,
            "client_secret": client_secret,
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "redirect_uris": [
                "https://hkpo-acceptance-generator-ucqvpysmj6jpdsb4a4knny.streamlit.app/",
                "http://localhost:8501/",
            ],
        }
    }

    # Detect current URL for redirect
    # On Streamlit Cloud the redirect must match exactly what was registered
    redirect_uri = "https://hkpo-acceptance-generator-ucqvpysmj6jpdsb4a4knny.streamlit.app/"

    flow = Flow.from_client_config(
        client_config,
        scopes=SCOPES,
        redirect_uri=redirect_uri,
    )
    return flow


def credentials_to_dict(creds: Credentials) -> dict:
    return {
        "token": creds.token,
        "refresh_token": creds.refresh_token,
        "token_uri": creds.token_uri,
        "client_id": creds.client_id,
        "client_secret": creds.client_secret,
        "scopes": creds.scopes,
    }


def dict_to_credentials(info: dict) -> Credentials:
    return Credentials(
        token=info["token"],
        refresh_token=info.get("refresh_token"),
        token_uri=info["token_uri"],
        client_id=info["client_id"],
        client_secret=info["client_secret"],
        scopes=info["scopes"],
    )


def get_drive_service(creds: Credentials):
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def find_folder_by_name(service, parent_id: str, folder_name: str):
    """Return the folder ID of the first match, or None."""
    query = (
        f"'{parent_id}' in parents and "
        f"name = '{folder_name}' and "
        f"mimeType = 'application/vnd.google-apps.folder' and "
        f"trashed = false"
    )
    results = service.files().list(
        q=query,
        spaces="drive",
        fields="files(id, name)",
        pageSize=10,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = results.get("files", [])
    if not files:
        return None
    return files[0]["id"]


def upload_file_to_folder(service, folder_id: str, filename: str, content: bytes, mime_type: str):
    file_metadata = {
        "name": filename,
        "parents": [folder_id],
    }
    media = MediaIoBaseUpload(BytesIO(content), mimetype=mime_type, resumable=True)
    uploaded = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id, name",
        supportsAllDrives=True,
    ).execute()
    return uploaded


def move_folder(service, folder_id: str, new_parent_id: str):
    # Get current parents
    file = service.files().get(
        fileId=folder_id,
        fields="parents",
        supportsAllDrives=True,
    ).execute()
    previous_parents = ",".join(file.get("parents", []))

    updated = service.files().update(
        fileId=folder_id,
        addParents=new_parent_id,
        removeParents=previous_parents,
        fields="id, parents",
        supportsAllDrives=True,
    ).execute()
    return updated


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="HKPO Acceptance Generator",
    page_icon="🎖️",
    layout="centered",
)

st.title("HKPO Veterans Academy")
st.subheader("Acceptance Letter & Training Plan Generator")
st.caption("Stage 2 – Generate documents + upload to Google Drive + move folder")

# ---- Authentication ----
if "credentials" not in st.session_state:
    st.session_state.credentials = None
if "oauth_code_verifier" not in st.session_state:
    st.session_state.oauth_code_verifier = None

# Handle OAuth callback
query_params = st.query_params
if "code" in query_params and st.session_state.credentials is None:
    try:
        flow = create_flow()
        # Prefer the verifier saved in session; fall back to the one passed in state
        verifier = st.session_state.oauth_code_verifier
        if not verifier and "state" in query_params:
            # state was used to carry the verifier
            verifier = query_params.get("state")
        if verifier:
            flow.code_verifier = verifier
        flow.fetch_token(code=query_params["code"])
        creds = flow.credentials
        st.session_state.credentials = credentials_to_dict(creds)
        st.session_state.oauth_code_verifier = None
        st.query_params.clear()
        st.rerun()
    except Exception as e:
        st.error(f"Login failed: {e}")
        st.info("Try clicking Login with Google again.")
        # Show debug info
        st.write("Debug – code_verifier present:", bool(st.session_state.oauth_code_verifier))
        st.stop()

# Login / Logout UI
if st.session_state.credentials is None:
    st.info("Please log in with the shared **director@hkpoveteransacademy.org** account to enable Google Drive actions.")
    if st.button("🔐 Login with Google", type="primary"):
        if not GOOGLE_AVAILABLE:
            st.error("Google libraries are missing. Check requirements.txt")
        else:
            try:
                flow = create_flow()
                # Generate verifier first so we control it
                import secrets
                code_verifier = secrets.token_urlsafe(64)
                flow.code_verifier = code_verifier
                st.session_state.oauth_code_verifier = code_verifier

                auth_url, _ = flow.authorization_url(
                    access_type="offline",
                    include_granted_scopes="true",
                    prompt="consent",
                    # Pass verifier in state as a backup in case session is lost
                    state=code_verifier,
                )
                st.write("Click the button below to authorize with Google:")
                st.link_button(
                    "Continue to Google Authorization",
                    auth_url,
                    type="primary",
                )
                st.caption("You will be sent to Google, then brought back here after you click Allow.")
            except Exception as e:
                st.error(f"Could not start login: {e}")
                st.info("Make sure Client ID and Client Secret are correctly set in Streamlit Secrets.")
    st.stop()

# Logged in
creds = dict_to_credentials(st.session_state.credentials)

col_a, col_b = st.columns([3, 1])
with col_a:
    st.success("Logged in with Google – Drive actions enabled")
with col_b:
    if st.button("Log out"):
        st.session_state.credentials = None
        st.rerun()

st.markdown("---")

# Sidebar status
with st.sidebar:
    st.header("Status")
    templates_ok = TEMPLATE_ROOT.exists() and ACCEPTANCE_TEMPLATE.exists()
    if templates_ok:
        st.success("Templates found")
    else:
        st.error("Templates missing")
        st.info("Make sure hkpo_templates/ is in the GitHub repo.")

# ---- Form ----
col1, col2 = st.columns(2)
with col1:
    first = st.text_input("Participant First Name *", placeholder="Jacob")
with col2:
    last = st.text_input("Participant Last Name *", placeholder="Weber")

co_director = st.text_input(
    "Co-Director Name (Training Plan signature)",
    value="Matthew Hale",
)

belt = st.selectbox(
    "Belt Level *",
    options=["Lean Practitioner", "Green Belt", "Black Belt"],
    index=1,
)

pmp = False
if belt != "Lean Practitioner":
    pmp = st.checkbox("Include PMP", value=False)

col3, col4 = st.columns(2)
with col3:
    start_date = st.date_input("Course Start Date *", value=date.today())
with col4:
    end_date = st.date_input("Course End Date *", value=date.today())

if end_date > start_date:
    days = (end_date - start_date).days + 1
    st.info(f"Calculated duration: **{days} days** (inclusive)")
else:
    st.warning("End date must be after start date")
    days = 0

st.markdown("---")

do_drive = st.checkbox(
    "Also upload to Google Drive and move folder to Accepted",
    value=True,
    help="Finds the folder named 'First Last' in Submissions, uploads the documents, then moves the folder to Accepted – Awaiting Start Date."
)

generate_btn = st.button("Generate Documents", type="primary", use_container_width=True)

if generate_btn:
    errors = []
    if not first.strip() or not last.strip():
        errors.append("First and Last name are required.")
    if not co_director.strip():
        errors.append("Co-Director name is required.")
    if end_date <= start_date:
        errors.append("End date must be after start date.")
    if not templates_ok:
        errors.append("Templates folder is missing.")

    if errors:
        for e in errors:
            st.error(e)
    else:
        with st.spinner("Generating documents…"):
            try:
                start_dt = datetime.combine(start_date, datetime.min.time())
                end_dt = datetime.combine(end_date, datetime.min.time())
                actual_days = inclusive_days(start_dt, end_dt)
                chosen_days, rel = find_closest_template(belt, pmp, actual_days)

                first_clean = first.strip()
                last_clean = last.strip()
                full_name = f"{first_clean} {last_clean}"

                acc_bytes = generate_acceptance(first_clean, last_clean, start_dt, end_dt)
                tp_bytes = generate_training_plan(
                    belt, pmp, actual_days,
                    first_clean, last_clean,
                    start_dt, end_dt, co_director.strip()
                )

                acc_name = f"{full_name} - HKPO VA Letter.docx"
                tp_name = f"{full_name} - HKPO VA Training Plan.docx"

                st.success("Documents generated successfully!")
                st.write(f"**Template used:** `{rel}` (closest to {actual_days} days)")

                # Download buttons
                st.markdown("### Download files")
                c1, c2 = st.columns(2)
                with c1:
                    st.download_button(
                        "Download Acceptance Letter",
                        acc_bytes,
                        file_name=acc_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
                with c2:
                    st.download_button(
                        "Download Training Plan",
                        tp_bytes,
                        file_name=tp_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )

                # ---- Drive actions ----
                if do_drive:
                    st.markdown("---")
                    st.subheader("Google Drive")
                    with st.spinner("Talking to Google Drive…"):
                        try:
                            service = get_drive_service(creds)

                            folder_id = find_folder_by_name(
                                service, SUBMISSIONS_FOLDER_ID, full_name
                            )

                            if not folder_id:
                                st.error(
                                    f"Could not find a folder named **“{full_name}”** "
                                    f"inside the Submissions folder.\n\n"
                                    "Check the exact spelling (including spaces and capitalization) "
                                    "and make sure the folder already exists in:\n"
                                    "`1. Student Folders → 1.2 Submissions - Working`"
                                )
                            else:
                                st.write(f"Found folder: **{full_name}**")

                                # Upload both files
                                upload_file_to_folder(
                                    service, folder_id, acc_name, acc_bytes,
                                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                )
                                st.write(f"✓ Uploaded `{acc_name}`")

                                upload_file_to_folder(
                                    service, folder_id, tp_name, tp_bytes,
                                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                )
                                st.write(f"✓ Uploaded `{tp_name}`")

                                # Move folder
                                move_folder(service, folder_id, ACCEPTED_FOLDER_ID)
                                st.success(
                                    f"Folder **“{full_name}”** moved to "
                                    f"**1.3 Accepted – Awaiting Start Date**"
                                )

                        except Exception as drive_err:
                            st.error(f"Drive error: {drive_err}")
                            st.info(
                                "Common causes:\n"
                                "- The login account does not have access to the Student Drive folders\n"
                                "- The folder name does not match exactly\n"
                                "- Drive API is not enabled in the Google Cloud project"
                            )

            except Exception as e:
                st.error(f"Error while generating documents:\n\n{e}")

st.markdown("---")
st.caption("HKPO Veterans Academy • Co-Director Continuity Tool • Stage 2 of 3")
