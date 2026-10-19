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
MOU_FOLDER_ID = "1DnmnDamSHJDBBlKVo1tWzLYkDgbJ7ebV"
# If start or end is after this date, do not attach MOU
MOU_CUTOFF = date(2027, 5, 7)


# OAuth scopes
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/gmail.compose",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/spreadsheets",
]

# Spreadsheet
SPREADSHEET_ID = "1opAbSoZE0fzejRyd-xm-nfdJOelBnL1gpQbrih3J9dE"
SHEET_NAME = "Participant Master"
HEADER_ROW = 5  # 1-indexed
NAME_COL_HEADER = "Name"
STATUS_COL_HEADER = "Manual Status"
EMAIL_COL_INDEX = 3  # Column D (0-based index)
ACCEPTED_STATUS = "Accepted by HKPO"

ACCEPTANCE_DRAFT_SUBJECT = "SkillBridge Acceptance - HKPO Veterans Academy"
ACCEPTANCE_CC = "jralston@hkpoveteransacademy.org"

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
    """Return (id, name) of exact folder match, or None."""
    # Escape single quotes in name for Drive query
    safe = folder_name.replace("'", "''")
    query = (
        f"'{parent_id}' in parents and "
        f"name = '{safe}' and "
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
    return files[0]["id"], files[0]["name"]


def list_submission_folders(service, parent_id: str):
    """Return list of {id, name} for all folders under Submissions."""
    folders = []
    page_token = None
    while True:
        results = service.files().list(
            q=(
                f"'{parent_id}' in parents and "
                f"mimeType = 'application/vnd.google-apps.folder' and "
                f"trashed = false"
            ),
            spaces="drive",
            fields="nextPageToken, files(id, name)",
            pageSize=200,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            pageToken=page_token,
        ).execute()
        for f in results.get("files", []):
            folders.append({"id": f["id"], "name": f["name"]})
        page_token = results.get("nextPageToken")
        if not page_token:
            break
    return folders


def suggest_folders(folders, query_name: str, limit: int = 3):
    """Return up to `limit` close folder name matches."""
    from difflib import SequenceMatcher
    needle = query_name.strip().lower()
    if not needle:
        return []
    scored = []
    for f in folders:
        name_l = f["name"].lower()
        ratio = SequenceMatcher(None, needle, name_l).ratio()
        tokens = [t for t in needle.split() if t]
        if tokens and all(t in name_l for t in tokens):
            ratio = max(ratio, 0.75)
        if needle in name_l or name_l in needle:
            ratio = max(ratio, 0.7)
        if ratio >= 0.45:
            scored.append((ratio, f))
    scored.sort(key=lambda x: -x[0])
    return [f for _, f in scored[:limit]]


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


def convert_docx_bytes_to_pdf(service, docx_bytes: bytes, base_name: str) -> bytes:
    """
    Convert DOCX to PDF using Google Drive:
    upload as Google Doc -> export PDF -> delete temp Google Doc.
    base_name is used for the temporary Google Doc name.
    """
    # Upload & convert to Google Doc
    file_metadata = {
        "name": f"_temp_{base_name}",
        "mimeType": "application/vnd.google-apps.document",
    }
    media = MediaIoBaseUpload(
        BytesIO(docx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        resumable=True,
    )
    gdoc = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id",
        supportsAllDrives=True,
    ).execute()
    gdoc_id = gdoc["id"]
    try:
        pdf_bytes = service.files().export(
            fileId=gdoc_id,
            mimeType="application/pdf",
        ).execute()
        if isinstance(pdf_bytes, str):
            pdf_bytes = pdf_bytes.encode("utf-8")
        return pdf_bytes
    finally:
        try:
            service.files().delete(fileId=gdoc_id, supportsAllDrives=True).execute()
        except Exception:
            pass


def find_mou_file(service):
    """Return (filename, file_id) of a PDF (or first non-image) in the MOU folder."""
    results = service.files().list(
        q=f"'{MOU_FOLDER_ID}' in parents and trashed = false and mimeType != 'application/vnd.google-apps.folder'",
        spaces="drive",
        fields="files(id, name, mimeType)",
        pageSize=20,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        orderBy="modifiedTime desc",
    ).execute()
    files = results.get("files", [])
    if not files:
        return None
    for f in files:
        name = f["name"].lower()
        mt = f.get("mimeType", "")
        if mt == "application/pdf" or name.endswith(".pdf"):
            return f["name"], f["id"]
    # fallback: first non-image
    for f in files:
        mt = f.get("mimeType", "")
        if not mt.startswith("image/"):
            return f["name"], f["id"]
    return None


def find_graphic_file(service):
    """Return (filename, file_id, mime_type) of an image in the MOU/assets folder."""
    results = service.files().list(
        q=f"'{MOU_FOLDER_ID}' in parents and trashed = false and mimeType != 'application/vnd.google-apps.folder'",
        spaces="drive",
        fields="files(id, name, mimeType)",
        pageSize=20,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        orderBy="modifiedTime desc",
    ).execute()
    files = results.get("files", [])
    image_exts = (".png", ".jpg", ".jpeg", ".gif", ".webp")
    for f in files:
        name = f["name"].lower()
        mt = f.get("mimeType", "")
        if mt.startswith("image/") or name.endswith(image_exts):
            if not mt.startswith("image/"):
                if name.endswith(".png"):
                    mt = "image/png"
                elif name.endswith((".jpg", ".jpeg")):
                    mt = "image/jpeg"
                elif name.endswith(".gif"):
                    mt = "image/gif"
                else:
                    mt = "image/png"
            return f["name"], f["id"], mt
    return None


def download_drive_file(service, file_id: str, mime_type_hint: str = None) -> bytes:
    """Download a binary file from Drive (or export Google Doc as PDF)."""
    meta = service.files().get(
        fileId=file_id,
        fields="mimeType, name",
        supportsAllDrives=True,
    ).execute()
    mt = meta.get("mimeType", "")
    if mt == "application/vnd.google-apps.document":
        data = service.files().export(fileId=file_id, mimeType="application/pdf").execute()
        return data if isinstance(data, bytes) else data.encode("utf-8")
    # Regular file
    from googleapiclient.http import MediaIoBaseDownload
    buf = BytesIO()
    request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return buf.read()


def needs_mou(start_dt: datetime, end_dt: datetime) -> bool:
    """False if start or end is after the MOU cutoff (May 7, 2027)."""
    cutoff = datetime(MOU_CUTOFF.year, MOU_CUTOFF.month, MOU_CUTOFF.day)
    return start_dt <= cutoff and end_dt <= cutoff



def move_folder(service, folder_id: str, new_parent_id: str):
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


def get_gmail_service(creds):
    return build("gmail", "v1", credentials=creds, cache_discovery=False)


def get_sheets_service(creds):
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def format_email_dates(start: datetime, end: datetime) -> str:
    """e.g. 12 October 2026 – 22 October 2027"""
    def part(d):
        return d.strftime("%d %B %Y").lstrip("0")
    return f"{part(start)} – {part(end)}"


def build_acceptance_email_html(first_name: str, dates_str: str, co_director: str, include_mou: bool = True, graphic_cid: str = None) -> str:
    mou_li = "<li>DoD SkillBridge Memorandum of Understanding (MOU)</li>" if include_mou else ""
    graphic_html = (
        f'<br><br><img src="cid:{graphic_cid}" alt="HKPO Veterans Academy" style="max-width:320px;height:auto;" />'
        if graphic_cid else ""
    )
    return f"""Dear {first_name},<br><br>
Congratulations! You have been officially accepted into the HKPO Veterans Academy. We are thrilled to welcome you into our upcoming Department of War SkillBridge cohort and look forward to supporting you throughout your transition.<br><br>
Your program dates are&nbsp;{dates_str}.<br><br>
Please find the following documents attached to support your SkillBridge application and coordination with your command:<br>
<ul>
<li>Official Acceptance Letter</li>
<li>Training Plan</li>
{mou_li}
</ul>
Kindly review these documents to ensure all details align with your requirements, including program dates.<br><br>
At this stage, we are awaiting your command’s approval. Once approved, please email a copy of your signed DoD SkillBridge authorization document to our team for our records. After receipt, we will proceed with the remaining onboarding steps.<br><br>
<strong>Next Steps</strong><br>
Please begin preparing and submit the following items to&nbsp;director@hkpoveteransacademy.org:<br>
<ul>
<li>A copy of your command-approved SkillBridge package (once available)</li>
<li>Point of contact information for your immediate supervisor and senior enlisted leader (name, rank, email, and phone number)</li>
<li>A photo of yourself in uniform (informal photos are acceptable and used for cohort materials)</li>
</ul>
<strong>Share Your Acceptance</strong><br>
We’ve also included an&nbsp;HKPO Veterans Academy acceptance badge/graphic&nbsp;for you to showcase your achievement on social media. Feel free to share and tag us—this helps expand awareness and support for transitioning service members.<br><br>
<strong>Onboarding Session</strong><br>
You will be scheduled for a brief&nbsp;30-minute onboarding session&nbsp;during the week prior to your start date, based on your availability. A Co-Director will reach out as your start date approaches to coordinate.<br><br>
If you have any questions, please don’t hesitate to reach out. We are excited to have you join the program and are committed to supporting you every step of the way.<br><br>
Warm regards,&nbsp;<br>
{co_director}<br>
Co-Director<br>
HKPO Veterans Academy<br>
A Nonprofit 501(c)3<br>
Support HKPO Veterans Academy&nbsp;<br>
Combined Federal Campaign (#43723)
{graphic_html}
"""


def build_acceptance_email_plain(first_name: str, dates_str: str, co_director: str, include_mou: bool = True) -> str:
    mou_line = "* DoD SkillBridge Memorandum of Understanding (MOU)\n" if include_mou else ""
    return f"""Dear {first_name},

Congratulations! You have been officially accepted into the HKPO Veterans Academy. We are thrilled to welcome you into our upcoming Department of War SkillBridge cohort and look forward to supporting you throughout your transition.

Your program dates are {dates_str}.

Please find the following documents attached to support your SkillBridge application and coordination with your command:
* Official Acceptance Letter
* Training Plan
{mou_line}
Kindly review these documents to ensure all details align with your requirements, including program dates.

At this stage, we are awaiting your command’s approval. Once approved, please email a copy of your signed DoD SkillBridge authorization document to our team for our records. After receipt, we will proceed with the remaining onboarding steps.

Next Steps
Please begin preparing and submit the following items to director@hkpoveteransacademy.org:
* A copy of your command-approved SkillBridge package (once available)
* Point of contact information for your immediate supervisor and senior enlisted leader (name, rank, email, and phone number)
* A photo of yourself in uniform (informal photos are acceptable and used for cohort materials)

Share Your Acceptance
We’ve also included an HKPO Veterans Academy acceptance badge/graphic for you to showcase your achievement on social media. Feel free to share and tag us—this helps expand awareness and support for transitioning service members.

Onboarding Session
You will be scheduled for a brief 30-minute onboarding session during the week prior to your start date, based on your availability. A Co-Director will reach out as your start date approaches to coordinate.

If you have any questions, please don’t hesitate to reach out. We are excited to have you join the program and are committed to supporting you every step of the way.

Warm regards, 
{co_director}
Co-Director
HKPO Veterans Academy
A Nonprofit 501(c)3
Support HKPO Veterans Academy 
Combined Federal Campaign (#43723)
"""


def create_message_with_attachments(
    to: str,
    subject: str,
    plain: str,
    html: str,
    attachments: list,
    cc: str = None,
    sender_name: str = "HKPO Veterans Academy",
    inline_images: list = None,
):
    """
    attachments = list of (filename, bytes, mime_type)
    inline_images = list of (cid, filename, bytes, mime_type)
    """
    import base64
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email.mime.image import MIMEImage
    from email import encoders

    msg = MIMEMultipart("mixed")
    msg["To"] = to
    msg["Subject"] = subject
    if cc:
        msg["Cc"] = cc
    msg["From"] = sender_name

    if inline_images:
        related = MIMEMultipart("related")
        alt = MIMEMultipart("alternative")
        alt.attach(MIMEText(plain, "plain"))
        alt.attach(MIMEText(html, "html"))
        related.attach(alt)
        for cid, filename, content, mime_type in inline_images:
            maintype, subtype = mime_type.split("/", 1)
            if maintype == "image":
                img = MIMEImage(content, _subtype=subtype)
            else:
                img = MIMEBase(maintype, subtype)
                img.set_payload(content)
                encoders.encode_base64(img)
            img.add_header("Content-ID", f"<{cid}>")
            img.add_header("Content-Disposition", "inline", filename=filename)
            related.attach(img)
        msg.attach(related)
    else:
        alt = MIMEMultipart("alternative")
        alt.attach(MIMEText(plain, "plain"))
        alt.attach(MIMEText(html, "html"))
        msg.attach(alt)

    for filename, content, mime_type in attachments:
        maintype, subtype = mime_type.split("/", 1)
        part = MIMEBase(maintype, subtype)
        part.set_payload(content)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", "attachment", filename=filename)
        msg.attach(part)

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return {"raw": raw}


def send_acceptance_email(
    gmail_service,
    to_email: str,
    first_name: str,
    dates_str: str,
    co_director: str,
    acc_pdf: bytes,
    acc_pdf_name: str,
    tp_pdf: bytes,
    tp_pdf_name: str,
    include_mou: bool = True,
    mou_bytes: bytes = None,
    mou_name: str = None,
    graphic_bytes: bytes = None,
    graphic_name: str = None,
    graphic_mime: str = None,
):
    subject = ACCEPTANCE_DRAFT_SUBJECT
    graphic_cid = "hkpo_badge" if graphic_bytes else None
    plain = build_acceptance_email_plain(first_name, dates_str, co_director, include_mou=include_mou)
    html = build_acceptance_email_html(
        first_name, dates_str, co_director,
        include_mou=include_mou,
        graphic_cid=graphic_cid,
    )
    attachments = [
        (acc_pdf_name, acc_pdf, "application/pdf"),
        (tp_pdf_name, tp_pdf, "application/pdf"),
    ]
    if include_mou and mou_bytes and mou_name:
        attachments.append((mou_name, mou_bytes, "application/pdf"))

    inline_images = None
    if graphic_bytes and graphic_name and graphic_mime:
        inline_images = [(graphic_cid, graphic_name, graphic_bytes, graphic_mime)]

    body = create_message_with_attachments(
        to=to_email,
        subject=subject,
        plain=plain,
        html=html,
        attachments=attachments,
        cc=ACCEPTANCE_CC,
        sender_name=f"{co_director} | HKPO Veterans Academy",
        inline_images=inline_images,
    )
    sent = gmail_service.users().messages().send(userId="me", body=body).execute()
    return sent



def _col_letter(idx0: int) -> str:
    n = idx0 + 1
    letters = ""
    while n:
        n, r = divmod(n - 1, 26)
        letters = chr(65 + r) + letters
    return letters


def load_participants(sheets_service):
    """Return list of dicts: {name, email, row_number, status} from Participant Master."""
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{SHEET_NAME}'",
    ).execute()
    values = result.get("values", [])
    if len(values) < HEADER_ROW:
        return []

    headers = [str(h).strip() for h in values[HEADER_ROW - 1]]
    try:
        name_idx = headers.index(NAME_COL_HEADER)
    except ValueError:
        return []

    try:
        status_idx = headers.index(STATUS_COL_HEADER)
    except ValueError:
        status_idx = None

    people = []
    for i in range(HEADER_ROW, len(values)):
        row = values[i]
        if name_idx >= len(row):
            continue
        name = str(row[name_idx]).strip()
        if not name:
            continue
        email = ""
        if EMAIL_COL_INDEX < len(row):
            email = str(row[EMAIL_COL_INDEX]).strip()
        status = ""
        if status_idx is not None and status_idx < len(row):
            status = str(row[status_idx]).strip()
        people.append({
            "name": name,
            "email": email,
            "row_number": i + 1,  # 1-based sheet row
            "status": status,
            "status_idx": status_idx,
        })
    return people


def find_exact_participant(people, full_name: str):
    for p in people:
        if p["name"] == full_name:
            return p
    return None


def suggest_participants(people, full_name: str, limit: int = 3):
    """Return up to `limit` close name matches (case-insensitive, token overlap / containment)."""
    from difflib import SequenceMatcher
    needle = full_name.strip().lower()
    if not needle:
        return []

    scored = []
    for p in people:
        name_l = p["name"].lower()
        ratio = SequenceMatcher(None, needle, name_l).ratio()
        # Boost if all tokens from the query appear in the name
        tokens = [t for t in needle.split() if t]
        if tokens and all(t in name_l for t in tokens):
            ratio = max(ratio, 0.75)
        if needle in name_l or name_l in needle:
            ratio = max(ratio, 0.7)
        if ratio >= 0.45:
            scored.append((ratio, p))
    scored.sort(key=lambda x: -x[0])
    return [p for _, p in scored[:limit]]


def set_status_on_row(sheets_service, row_number: int, status_idx: int) -> None:
    col = _col_letter(status_idx)
    cell = f"'{SHEET_NAME}'!{col}{row_number}"
    sheets_service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=cell,
        valueInputOption="USER_ENTERED",
        body={"values": [[ACCEPTED_STATUS]]},
    ).execute()


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
st.caption("Stage 3 – Generate + Drive + Email + Status update")

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
    "Co-Director Name (Training Plan signature & email)",
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
    "Upload to Google Drive and move folder to Accepted",
    value=True,
    help="Finds the folder named 'First Last' in Submissions, uploads the documents, then moves the folder to Accepted – Awaiting Start Date."
)
do_email = st.checkbox(
    "Send acceptance email",
    value=True,
    help="Sends the acceptance email with Acceptance Letter + Training Plan attached. MOU should already be on your Gmail draft workflow; this composes the personalized message.",
)
do_status = st.checkbox(
    "Update spreadsheet status to Accepted by HKPO",
    value=True,
)


# ---- Folder / sheet match confirmation (survives rerun) ----
# Folder suggestions first
if st.session_state.get("folder_suggestions"):
    st.warning(st.session_state.get("folder_suggest_msg", "Exact folder name not found."))
    st.write("**Drive folder** – Did you mean one of these? Click to confirm:")
    for i, f in enumerate(st.session_state["folder_suggestions"]):
        if st.button(f"📁 {f['name']}", key=f"pick_folder_{i}_{f['id']}"):
            query = st.session_state.get("pending_job", {}).get("form_name", "")
            st.session_state.selected_folder = {**f, "query_name": query}
            st.session_state.folder_suggestions = None
            st.session_state.folder_suggest_msg = None
            st.session_state.auto_run = True
            st.rerun()
    st.caption("After you click a folder, processing will continue (sheet match may come next).")
    st.stop()

# Sheet suggestions second
if st.session_state.get("sheet_suggestions"):
    st.warning(st.session_state.get("sheet_suggest_msg", "Exact spreadsheet name not found."))
    st.write("**Spreadsheet row** – Did you mean one of these? Click to confirm:")
    for i, s in enumerate(st.session_state["sheet_suggestions"]):
        label = f"{s['name']}"
        if s.get("email"):
            label += f"  ·  {s['email']}"
        if s.get("status"):
            label += f"  ·  [{s['status']}]"
        if st.button(label, key=f"pick_sheet_{i}_{s['row_number']}"):
            query = st.session_state.get("pending_job", {}).get("form_name", "")
            st.session_state.selected_participant = {**s, "query_name": query}
            st.session_state.sheet_suggestions = None
            st.session_state.sheet_suggest_msg = None
            st.session_state.auto_run = True
            st.rerun()
    st.caption("After you click a row, processing will continue automatically.")
    st.stop()

if st.session_state.get("selected_folder"):
    st.success(f"Selected Drive folder: **{st.session_state.selected_folder['name']}**")

if st.session_state.get("selected_participant"):
    sel = st.session_state.selected_participant
    st.success(
        f"Selected spreadsheet row: **{sel['name']}**"
        + (f" · {sel.get('email','')}" if sel.get("email") else "")
    )

generate_btn = st.button("Generate & Process", type="primary", use_container_width=True)

if generate_btn or st.session_state.get("auto_run"):
    st.session_state.auto_run = False
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
                # Form spelling is always used for documents + email body
                form_name = f"{first_clean} {last_clean}"

                st.session_state.pending_job = {
                    "form_name": form_name,
                    "first": first_clean,
                    "last": last_clean,
                    "co_director": co_director.strip(),
                    "belt": belt,
                    "pmp": pmp,
                    "start": start_date.isoformat(),
                    "end": end_date.isoformat(),
                    "do_drive": do_drive,
                    "do_email": do_email,
                    "do_status": do_status,
                }

                # Clear stale selections if the typed name changed
                sf = st.session_state.get("selected_folder")
                if sf and sf.get("query_name") != form_name:
                    st.session_state.selected_folder = None
                    sf = None
                sp = st.session_state.get("selected_participant")
                if sp and sp.get("query_name") != form_name:
                    st.session_state.selected_participant = None
                    sp = None

                service = get_drive_service(creds) if (do_drive or do_email) else None

                # ---- Drive folder match (exact, then suggest) ----
                if do_drive and not sf:
                    try:
                        exact_folder = find_folder_by_name(service, SUBMISSIONS_FOLDER_ID, form_name)
                        if exact_folder:
                            folder_id, folder_name = exact_folder
                            sf = {"id": folder_id, "name": folder_name, "query_name": form_name}
                            st.session_state.selected_folder = sf
                        else:
                            folders = list_submission_folders(service, SUBMISSIONS_FOLDER_ID)
                            suggestions = suggest_folders(folders, form_name, limit=3)
                            if not suggestions:
                                st.error(
                                    f"No folder named **{form_name}** in Submissions, and no similar "
                                    "folder names were found."
                                )
                                st.stop()
                            st.session_state.folder_suggestions = suggestions
                            st.session_state.folder_suggest_msg = (
                                f"Exact folder name not found for **{form_name}** in Submissions."
                            )
                            st.rerun()
                    except Exception as folder_err:
                        st.error(f"Could not look up Drive folders: {folder_err}")
                        st.stop()

                # ---- Spreadsheet row match (exact, then suggest) ----
                if (do_email or do_status) and not sp:
                    try:
                        sheets = get_sheets_service(creds)
                        people = load_participants(sheets)
                        exact = find_exact_participant(people, form_name)
                        if exact:
                            sp = {**exact, "query_name": form_name}
                            st.session_state.selected_participant = sp
                        else:
                            suggestions = suggest_participants(people, form_name, limit=3)
                            if not suggestions:
                                st.error(
                                    f"Exact match not found for **{form_name}** in Participant Master, "
                                    "and no similar names were found."
                                )
                                st.stop()
                            st.session_state.sheet_suggestions = suggestions
                            st.session_state.sheet_suggest_msg = (
                                f"Exact spreadsheet name not found for **{form_name}**."
                            )
                            st.rerun()
                    except Exception as lookup_err:
                        st.error(f"Could not look up participants: {lookup_err}")
                        st.info("Enable Sheets API and re-login if needed.")
                        st.stop()

                selected = sp  # spreadsheet row (email + status)
                selected_folder = sf  # drive folder

                if selected_folder:
                    st.info(f"Drive folder: **{selected_folder['name']}**")
                if selected:
                    st.info(
                        f"Spreadsheet row: **{selected['name']}**"
                        + (f" · {selected['email']}" if selected.get("email") else "")
                    )
                    if (do_email or do_status) and not selected.get("email"):
                        st.error(
                            f"No email found in column D for **{selected['name']}**. "
                            "Add their email to the sheet and try again."
                        )
                        st.stop()

                # Documents + email body always use form spelling
                full_name = form_name


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

                # PDF conversion + Drive + Email prep (needs Drive service)
                acc_pdf = tp_pdf = None
                acc_pdf_name = f"{full_name} - HKPO VA Letter.pdf"
                tp_pdf_name = f"{full_name} - HKPO VA Training Plan.pdf"
                include_mou = needs_mou(start_dt, end_dt)
                mou_bytes = mou_name = None

                if do_drive or do_email:
                    st.markdown("---")
                    st.subheader("PDF conversion")
                    with st.spinner("Converting documents to PDF via Google Drive…"):
                        try:
                            service = get_drive_service(creds)
                            acc_pdf = convert_docx_bytes_to_pdf(service, acc_bytes, acc_name)
                            tp_pdf = convert_docx_bytes_to_pdf(service, tp_bytes, tp_name)
                            st.success("PDF conversion complete")
                        except Exception as pdf_err:
                            st.error(f"PDF conversion error: {pdf_err}")
                            st.info("Drive API must be enabled and the login account needs Drive access.")
                            acc_pdf = tp_pdf = None

                    if not include_mou:
                        st.warning(
                            f"**MOU notice:** Program dates for **{full_name}** extend past "
                            f"May 7, 2027. The MOU was **not** attached. "
                            f"Please provide this applicant with a copy of the next MOU."
                        )

                if do_drive:
                    st.markdown("---")
                    st.subheader("Google Drive")
                    with st.spinner("Talking to Google Drive…"):
                        try:
                            service = get_drive_service(creds)

                            if not selected_folder:
                                st.error("No Drive folder selected.")
                            else:
                                folder_id = selected_folder["id"]
                                st.write(f"Using folder: **{selected_folder['name']}**")

                                # Upload DOCX
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

                                # Upload PDF if converted
                                if acc_pdf:
                                    upload_file_to_folder(
                                        service, folder_id, acc_pdf_name, acc_pdf, "application/pdf"
                                    )
                                    st.write(f"✓ Uploaded `{acc_pdf_name}`")
                                if tp_pdf:
                                    upload_file_to_folder(
                                        service, folder_id, tp_pdf_name, tp_pdf, "application/pdf"
                                    )
                                    st.write(f"✓ Uploaded `{tp_pdf_name}`")

                                # Move folder
                                move_folder(service, folder_id, ACCEPTED_FOLDER_ID)
                                st.success(
                                    f"Folder **“{selected_folder['name']}”** moved to "
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

                # ---- Stage 3: Email (PDF only + optional MOU) ----
                if do_email:
                    st.markdown("---")
                    st.subheader("Acceptance Email")
                    if not acc_pdf or not tp_pdf:
                        st.error("Cannot send email – PDF conversion failed. Fix PDF conversion first.")
                    else:
                        with st.spinner("Sending acceptance email…"):
                            try:
                                service = get_drive_service(creds)
                                mou_bytes = mou_name = None
                                if include_mou:
                                    mou_info = find_mou_file(service)
                                    if mou_info:
                                        mou_name, mou_id = mou_info
                                        mou_bytes = download_drive_file(service, mou_id)
                                        if not mou_name.lower().endswith(".pdf"):
                                            mou_name = mou_name.rsplit(".", 1)[0] + ".pdf"
                                        st.write(f"Attached MOU: `{mou_name}`")
                                    else:
                                        st.warning("No MOU PDF found in the assets folder – sending without MOU.")
                                        include_mou = False

                                graphic_bytes = graphic_name = graphic_mime = None
                                graphic_info = find_graphic_file(service)
                                if graphic_info:
                                    graphic_name, graphic_id, graphic_mime = graphic_info
                                    graphic_bytes = download_drive_file(service, graphic_id)
                                    st.write(f"Embedded graphic: `{graphic_name}`")
                                else:
                                    st.warning("No image graphic found in the assets folder – email will send without the badge.")

                                gmail = get_gmail_service(creds)
                                dates_str = format_email_dates(start_dt, end_dt)
                                send_acceptance_email(
                                    gmail,
                                    to_email=selected['email'],
                                    first_name=first_clean,
                                    dates_str=dates_str,
                                    co_director=co_director.strip(),
                                    acc_pdf=acc_pdf,
                                    acc_pdf_name=acc_pdf_name,
                                    tp_pdf=tp_pdf,
                                    tp_pdf_name=tp_pdf_name,
                                    include_mou=include_mou,
                                    mou_bytes=mou_bytes,
                                    mou_name=mou_name,
                                    graphic_bytes=graphic_bytes,
                                    graphic_name=graphic_name,
                                    graphic_mime=graphic_mime,
                                )
                                st.success(
                                    f"Acceptance email sent to **{selected['email']}** "
                                    f"(CC: {ACCEPTANCE_CC}) with PDF attachments"
                                    + (" + MOU" if include_mou else " (no MOU)")
                                )
                            except Exception as mail_err:
                                st.error(f"Email error: {mail_err}")
                                st.info(
                                    "Enable Gmail API in Google Cloud if needed, and re-login "
                                    "so the new Gmail permission is granted."
                                )

                # ---- Stage 3: Spreadsheet status ----
                if do_status and selected:
                    st.markdown("---")
                    st.subheader("Spreadsheet Status")
                    with st.spinner("Updating Participant Master…"):
                        try:
                            sheets = get_sheets_service(creds)
                            if selected.get("status_idx") is None:
                                st.error("Could not find Manual Status column.")
                            else:
                                set_status_on_row(
                                    sheets,
                                    selected["row_number"],
                                    selected["status_idx"],
                                )
                                st.success(
                                    f"Manual Status for **{selected['name']}** "
                                    f"(row {selected['row_number']}) set to "
                                    f"**Accepted by HKPO**"
                                )
                                st.session_state.selected_participant = None
                                st.session_state.selected_folder = None
                        except Exception as sheet_err:
                            st.error(f"Sheets error: {sheet_err}")
                            st.info(
                                "Enable Google Sheets API in Google Cloud if needed, and "
                                "re-login so the Sheets permission is granted."
                            )

            except Exception as e:
                st.error(f"Error while generating documents:\n\n{e}")

st.markdown("---")
st.caption("HKPO Veterans Academy • Co-Director Continuity Tool • Stage 3")
