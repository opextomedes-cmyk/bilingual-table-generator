# 📄 Bilingual Table Generator
**Tomedes OPEX DTP Team**

A web app that extracts text from a PDF and places it into a formatted bilingual Word table — powered by Claude AI.

---

## 🚀 Deploy in 5 Minutes (Free)

### Step 1 — Fork or upload this repo to GitHub
- Go to github.com → New repository → Name it `bilingual-table-generator`
- Upload all files from this folder

### Step 2 — Deploy on Streamlit Cloud (Free)
1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Sign in with GitHub
3. Click **New app**
4. Select your repository
5. Set **Main file path** to `app.py`
6. Click **Deploy**

Your app will be live at:
`https://your-username-bilingual-table-generator-app-XXXXX.streamlit.app`

Share that link with your team!

---

## 📋 How to Use

1. Enter your **Claude API Key** (get one free at console.anthropic.com)
2. Upload your **source PDF**
3. Upload your **Word template** (.docx)
4. Set source & target language
5. Click **Generate Bilingual Table**
6. Download the `.docx` file

---

## 🔧 Requirements

All handled automatically by Streamlit Cloud:
- streamlit
- pymupdf
- python-docx
- anthropic

---

## 📝 Rules Applied
- Text extracted directly — no transcription
- Bold, italic, underline preserved
- Organized by page with Page X headers
- Non-extractable content flagged automatically
- Calibri font at selected size
- Target column blank and ready for translation
