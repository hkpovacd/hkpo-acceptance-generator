# HKPO Acceptance Letter & Training Plan Generator
**Stage 1 – Document Generation (Streamlit Cloud Ready)**

This version is prepared for **Streamlit Community Cloud** so both Co-Directors can use it from a normal web browser with zero local installation.

---

## Deploy to Streamlit Community Cloud (Recommended)

### 1. Put the templates in the project
Copy your template files into the `hkpo_templates` folder exactly like this:

```
hkpo_templates/
├── acceptance/
│   └── acceptance_letter.docx
├── lean_practitioner/
│   └── 60_day.docx
├── green_belt/
│   ├── 60_day.docx
│   ├── 63_day.docx
│   ├── 88_day.docx
│   └── 90_day.docx
├── green_belt_pmp/
│   ├── 60_day.docx
│   ├── 73_day.docx
│   └── 92_day.docx
├── black_belt/
│   ├── 90_day.docx
│   └── 120_day.docx
└── black_belt_pmp/
    ├── 89_day.docx
    ├── 90_day.docx
    ├── 105_day.docx
    ├── 120_day.docx
    └── 180_day.docx
```

### 2. Create a GitHub repository
1. Go to https://github.com and sign in (or create a free account).
2. Click **New repository**.
3. Name it something like `hkpo-acceptance-generator`.
4. Set it to **Private** (recommended because it contains your templates).
5. Click **Create repository**.
6. Upload all the files from the `hkpo_acceptance_app` folder (including the `hkpo_templates` folder with your documents).

### 3. Deploy on Streamlit Community Cloud
1. Go to https://share.streamlit.io
2. Sign in with the **same GitHub account**.
3. Click **New app**.
4. Select the repository you just created.
5. Main file path: `app.py`
6. Click **Deploy**.

After a minute or two you will get a URL such as:  
`https://hkpo-acceptance-generator.streamlit.app`

Both Co-Directors can now just open that link.

---

## Important notes for the cloud version

| Topic | Details |
|-------|---------|
| **PDF files** | Streamlit Cloud free tier does not include LibreOffice, so only `.docx` files are generated. They open perfectly in Word / Google Docs. We can add PDF conversion later using Google Docs API if needed. |
| **Templates** | Must be inside the GitHub repository (as shown above). |
| **Login** | For Stage 1 no login is required. Stage 2 & 3 will use the shared `director@hkpoveteransacademy.org` account. |
| **Updating the app** | Just push changes to the GitHub repo — Streamlit Cloud automatically redeploys. |

---

## Local testing (optional)
If you still want to test on your own computer:

```bash
pip install -r requirements.txt
streamlit run app.py
```

---

## Folder IDs ready for Stage 2
- Submissions: `12MfLDc8Jb4m6b5ke6NaEdiN433SQ0MRV`
- Accepted: `1SQ_yDzpWCdzKfl6K_QA2UQWklCjipLYz`

---

**Next stages**
- Stage 2 → Automatic Google Drive upload + folder move  
- Stage 3 → Acceptance email + status update to “Accepted by HKPO”
