#!/usr/bin/env python3
"""
HKPO Veterans Academy – Acceptance Letter & Training Plan Generator
Streamlit Web Version (Stage 1)
"""

import os
import re
import shutil
import subprocess
import tempfile
from datetime import datetime, date
from pathlib import Path
from copy import deepcopy

import streamlit as st

try:
    from docx import Document
except ImportError:
    st.error("Missing dependency. Run: `pip install python-docx`")
    st.stop()

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
TEMPLATE_ROOT = SCRIPT_DIR / "hkpo_templates"
OUTPUT_DIR = SCRIPT_DIR / "output"

# Available templates: (duration_days, filename)
TEMPLATES = {
    "Lean Practitioner": {
        False: [  # no PMP
            (60, "lean_practitioner/60_day.docx"),
        ],
        True: []  # Lean Practitioner has no PMP variant
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
# Helpers (same logic as your original script)
# ---------------------------------------------------------------------------
def find_closest_template(belt: str, pmp: bool, approx_days: int):
    options = TEMPLATES.get(belt, {}).get(pmp, [])
    if not options:
        raise ValueError(
            f"No templates available for {belt} {'+ PMP' if pmp else '(no PMP)'}."
        )
    best = min(options, key=lambda x: abs(x[0] - approx_days))
    return best


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


def _find_libreoffice() -> str | None:
    import sys
    for name in ("soffice", "libreoffice", "soffice.exe"):
        if shutil.which(name):
            return name
    if sys.platform.startswith("win"):
        candidates = [
            Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
            Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
        ]
        for p in candidates:
            if p.exists():
                return str(p)
    mac_path = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    if mac_path.exists():
        return str(mac_path)
    return None


def convert_to_pdf(docx_path: Path, pdf_path: Path) -> bool:
    lo = _find_libreoffice()
    if not lo:
        return False
    out_dir = pdf_path.parent
    try:
        cmd = [
            lo, "--headless", "--convert-to", "pdf",
            "--outdir", str(out_dir), str(docx_path),
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=90)
        generated = out_dir / (docx_path.stem + ".pdf")
        if generated.exists() and generated != pdf_path:
            generated.rename(pdf_path)
        return pdf_path.exists()
    except Exception:
        return False


def generate_acceptance(first: str, last: str, start: datetime, end: datetime,
                        out_dir: Path) -> tuple[Path, Path]:
    if not ACCEPTANCE_TEMPLATE.exists():
        raise FileNotFoundError(
            f"Acceptance letter template not found:\n{ACCEPTANCE_TEMPLATE}\n\n"
            "Please place the template file at that location."
        )
    doc = Document(ACCEPTANCE_TEMPLATE)
    full_name = f"{first} {last}"
    date_str = format_acceptance_dates(start, end)

    replacements = {
        "First Last": full_name,
        "Month XX, 20XX through Month XX, 20XX": date_str,
    }
    replace_everywhere(doc, replacements)

    base = f"{first} {last} - HKPO VA Letter"
    docx_path = out_dir / f"{base}.docx"
    pdf_path = out_dir / f"{base}.pdf"
    doc.save(docx_path)
    convert_to_pdf(docx_path, pdf_path)
    return docx_path, pdf_path


def generate_training_plan(belt: str, pmp: bool, approx_days: int,
                           first: str, last: str,
                           start: datetime, end: datetime,
                           co_director: str,
                           out_dir: Path) -> tuple[Path, Path, int]:
    chosen_days, rel_path = find_closest_template(belt, pmp, approx_days)
    src = TEMPLATE_ROOT / rel_path
    if not src.exists():
        raise FileNotFoundError(f"Template not found: {src}")

    doc = Document(src)
    full_name = f"{first} {last}"
    date_str = f"{format_date_long(start)} – {format_date_long(end)}"
    actual_days = inclusive_days(start, end)

    # Greeting
    for p in doc.paragraphs:
        if p.text.strip().startswith("Dear "):
            for run in p.runs:
                run.text = ""
            if p.runs:
                p.runs[0].text = f"Dear {full_name},"
            else:
                p.add_run(f"Dear {full_name},")
            break

    # Course Dates
    for p in doc.paragraphs:
        txt = p.text
        if "Course Dates" in txt or "Course dates" in txt:
            new = re.sub(
                r"(Course Dates[:\s]*)(.*)",
                rf"\1{date_str}",
                txt,
                flags=re.IGNORECASE,
            )
            for run in p.runs:
                run.text = ""
            if p.runs:
                p.runs[0].text = new
            else:
                p.add_run(new)

    # Length of Training
    for p in doc.paragraphs:
        txt = p.text
        if "Length of Training" in txt:
            new = re.sub(
                r"(Length of Training[:\s]*)(.*)",
                rf"\g<1>{actual_days} days",
                txt,
                flags=re.IGNORECASE,
            )
            for run in p.runs:
                run.text = ""
            if p.runs:
                p.runs[0].text = new
            else:
                p.add_run(new)

    # Co-Director name
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

    base = f"{first} {last} - HKPO VA Training Plan"
    docx_path = out_dir / f"{base}.docx"
    pdf_path = out_dir / f"{base}.pdf"
    doc.save(docx_path)
    convert_to_pdf(docx_path, pdf_path)
    return docx_path, pdf_path, actual_days


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
st.caption("Stage 1 – Document generation (Drive upload & email coming next)")

# Sidebar status
with st.sidebar:
    st.header("Status")
    templates_ok = TEMPLATE_ROOT.exists() and ACCEPTANCE_TEMPLATE.exists()
    if templates_ok:
        st.success("Templates folder found")
    else:
        st.error("Templates folder missing")
        st.info(
            "Place your templates in the `hkpo_templates` folder "
            "inside this project (see README for exact structure)."
        )

    lo = _find_libreoffice()
    if lo:
        st.success("LibreOffice found – PDF conversion available")
    else:
        st.info("Running in cloud mode – generating .docx files only "
                "(fully usable in Word / Google Docs). PDF can be added later.")

st.markdown("---")

# ---- Form ----
col1, col2 = st.columns(2)
with col1:
    first = st.text_input("Participant First Name *", placeholder="Jacob")
with col2:
    last = st.text_input("Participant Last Name *", placeholder="Weber")

co_director = st.text_input(
    "Co-Director Name (appears on Training Plan & email signature)",
    value="Matthew Hale",
    help="This name will be used on the Training Plan signature and in the acceptance email."
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

# Live duration
if end_date > start_date:
    days = (end_date - start_date).days + 1
    st.info(f"Calculated duration: **{days} days** (inclusive)")
else:
    st.warning("End date must be after start date")
    days = 0

st.markdown("---")

generate_btn = st.button("Generate Documents", type="primary", use_container_width=True)

if generate_btn:
    # Validation
    errors = []
    if not first.strip() or not last.strip():
        errors.append("First and Last name are required.")
    if not co_director.strip():
        errors.append("Co-Director name is required.")
    if end_date <= start_date:
        errors.append("End date must be after start date.")
    if not templates_ok:
        errors.append("Templates folder is missing. See sidebar for instructions.")

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

                # Create output folder
                person_dir = OUTPUT_DIR / f"{first.strip()} {last.strip()}"
                person_dir.mkdir(parents=True, exist_ok=True)

                # Generate
                acc_docx, acc_pdf = generate_acceptance(
                    first.strip(), last.strip(), start_dt, end_dt, person_dir
                )
                tp_docx, tp_pdf, used_days = generate_training_plan(
                    belt, pmp, actual_days,
                    first.strip(), last.strip(),
                    start_dt, end_dt, co_director.strip(),
                    person_dir
                )

                st.success("Documents generated successfully!")
                st.write(f"**Template used:** `{rel}` (closest to {actual_days} days)")
                st.write(f"**Output folder:** `{person_dir}`")

                # Download buttons
                st.markdown("### Download files")
                col_a, col_b = st.columns(2)

                with col_a:
                    st.markdown("**Acceptance Letter**")
                    with open(acc_docx, "rb") as f:
                        st.download_button(
                            "Download .docx",
                            f,
                            file_name=acc_docx.name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        )
                    if acc_pdf.exists():
                        with open(acc_pdf, "rb") as f:
                            st.download_button(
                                "Download .pdf",
                                f,
                                file_name=acc_pdf.name,
                                mime="application/pdf",
                            )
                    else:
                        st.caption("PDF not available in cloud mode (docx works in Word/Google Docs)")

                with col_b:
                    st.markdown("**Training Plan**")
                    with open(tp_docx, "rb") as f:
                        st.download_button(
                            "Download .docx",
                            f,
                            file_name=tp_docx.name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        )
                    if tp_pdf.exists():
                        with open(tp_pdf, "rb") as f:
                            st.download_button(
                                "Download .pdf",
                                f,
                                file_name=tp_pdf.name,
                                mime="application/pdf",
                            )
                    else:
                        st.caption("PDF not available in cloud mode (docx works in Word/Google Docs)")

                st.info(
                    "Next stage will automatically upload these files to the applicant’s "
                    "Google Drive folder and move it to “Accepted – Awaiting Start Date”."
                )

            except Exception as e:
                st.error(f"Error while generating documents:\n\n{e}")

# Footer
st.markdown("---")
st.caption(
    "HKPO Veterans Academy • Co-Director Continuity Tool • Stage 1 of 3"
)
