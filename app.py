"""
=============================================================
  BILINGUAL TABLE GENERATOR — Tomedes OPEX DTP Team
  Powered by Claude AI + Streamlit
=============================================================
"""

import streamlit as st
import fitz  # PyMuPDF
import anthropic
import json
import re
import io
import os
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── Page config ───────────────────────────────────────────────
st.set_page_config(
    page_title="Bilingual Table Generator",
    page_icon="📄",
    layout="centered"
)

# ── Styling ───────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #f0f4f8; }
    .stApp { max-width: 900px; margin: 0 auto; }
    h1 { color: #1E2761 !important; }
    h2, h3 { color: #1E2761 !important; }
    .stButton > button {
        background-color: #1E2761;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 0.5rem 2rem;
        border: none;
        width: 100%;
    }
    .stButton > button:hover {
        background-color: #2d3a8a;
        color: white;
    }
    .tip-box {
        background: #fff8e1;
        border-left: 4px solid #ffc107;
        padding: 12px 16px;
        border-radius: 6px;
        font-size: 13px;
        color: #5d4037;
        margin: 10px 0;
    }
    .success-box {
        background: #f0fff4;
        border: 2px solid #a5d6a7;
        padding: 16px;
        border-radius: 10px;
        text-align: center;
        margin-top: 16px;
    }
    .header-box {
        background: #1E2761;
        color: white;
        padding: 20px 24px;
        border-radius: 12px;
        margin-bottom: 24px;
    }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────
st.markdown("""
<div class="header-box">
    <h2 style="color:white;margin:0;">📄 Bilingual Table Generator</h2>
    <p style="color:#CADCFC;margin:4px 0 0;">Tomedes OPEX DTP Team — Upload PDF + Template → Get Bilingual .docx</p>
</div>
""", unsafe_allow_html=True)


# ── XML helpers for DOCX ──────────────────────────────────────
def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_cell_borders(cell, color="B0BEC5", size="4"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ('top', 'left', 'bottom', 'right'):
        border = OxmlElement(f'w:{side}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), size)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), color)
        tcBorders.append(border)
    tcPr.append(tcBorders)

def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    mar = OxmlElement('w:tcMar')
    for side, val in [('top', top), ('left', left), ('bottom', bottom), ('right', right)]:
        m = OxmlElement(f'w:{side}')
        m.set(qn('w:w'), str(val))
        m.set(qn('w:type'), 'dxa')
        mar.append(m)
    tcPr.append(mar)

def set_col_width(cell, width_dxa):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(width_dxa))
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)


# ── PDF → base64 for Claude ───────────────────────────────────
import base64

def pdf_to_base64(pdf_bytes):
    return base64.b64encode(pdf_bytes).decode('utf-8')


# ── Claude extraction ─────────────────────────────────────────
def extract_with_claude(pdf_bytes, api_key):
    client = anthropic.Anthropic(api_key=api_key)
    pdf_b64 = pdf_to_base64(pdf_bytes)

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=8000,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": pdf_b64
                    }
                },
                {
                    "type": "text",
                    "text": """Extract ALL text from this PDF. Return ONLY raw JSON — no markdown, no explanation, no code fences.

CRITICAL RULES:
- DO NOT transcribe, retype, paraphrase, or summarize — extract EXACTLY as written
- Preserve bold, italic, underline indicators
- Keep sentences COMPLETE — never break mid-sentence
- Top-to-bottom, left-to-right reading order per page
- Non-text elements (stamps, images, seals, signatures): use type "non_extractable"
- Each paragraph / heading / bullet / numbered item = one separate element
- Include ALL text: headers, footers, labels, captions, field names, field values

Return this exact JSON format:
{"pages":[{"pageNumber":1,"elements":[{"type":"heading|paragraph|bullet|numbered|caption|label|non_extractable","text":"exact extracted text","bold":false,"italic":false,"underline":false}]}],"totalPages":1}"""
                }
            ]
        }]
    )

    raw = response.content[0].text if response.content else ""
    raw = raw.replace("```json", "").replace("```", "").strip()

    try:
        return json.loads(raw)
    except Exception:
        match = re.search(r'\{[\s\S]*\}', raw)
        if match:
            try:
                return json.loads(match.group(0))
            except Exception:
                pass
    return {"pages": [], "totalPages": 0}


# ── Build DOCX ────────────────────────────────────────────────
def build_docx(pages, source_lang, target_lang, font_size):
    doc = Document()

    # Page setup
    section = doc.sections[0]
    section.page_width  = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin   = Cm(2)
    section.right_margin  = Cm(2)
    section.top_margin    = Cm(2)
    section.bottom_margin = Cm(2)

    # Default style
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(font_size)

    COL_W = 4050  # DXA

    NAVY    = "1E2761"
    ICE     = "D9E2F3"
    BLUE    = "2E74B5"
    WHITE   = "FFFFFF"
    SHADE   = "F4F6FB"
    RED     = "C62828"

    def make_header_row(label):
        row = doc.add_table(rows=1, cols=2).rows[0]  # temp — we'll use table below
        pass

    # Build one continuous table
    table = doc.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    def add_col_header():
        row = table.add_row()
        for i, lang_label in enumerate([source_lang, target_lang]):
            c = row.cells[i]
            set_cell_bg(c, BLUE)
            set_cell_borders(c, color=BLUE, size="6")
            set_cell_margins(c, top=100, bottom=100, left=150, right=150)
            set_col_width(c, COL_W)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(lang_label)
            r.bold = True
            r.font.name = 'Calibri'
            r.font.size = Pt(font_size)
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    def add_page_header(label):
        row = table.add_row()
        for i in range(2):
            c = row.cells[i]
            set_cell_bg(c, ICE)
            set_cell_borders(c, color="1E2761", size="6")
            set_cell_margins(c, top=100, bottom=100, left=150, right=150)
            set_col_width(c, COL_W)
        p = row.cells[0].paragraphs[0]
        r = p.add_run(label)
        r.bold = True
        r.font.name = 'Calibri'
        r.font.size = Pt(font_size)
        r.font.color.rgb = RGBColor(0x1E, 0x27, 0x61)

    def add_content_row(el, shade, page_num):
        row = table.add_row()
        src = row.cells[0]
        tgt = row.cells[1]
        bg = SHADE if shade else WHITE

        for c in [src, tgt]:
            set_cell_bg(c, bg)
            set_cell_borders(c, color="B0BEC5", size="2")
            set_cell_margins(c)
            set_col_width(c, COL_W)

        # Source text
        p = src.paragraphs[0]
        txt = el.get("text", "")
        color_rgb = None

        if el.get("type") == "non_extractable":
            txt = f"[NON-EXTRACTABLE CONTENT — Page {page_num}]"
            r = p.add_run(txt)
            r.italic = True
            r.font.name = 'Calibri'
            r.font.size = Pt(font_size)
            r.font.color.rgb = RGBColor(0xC6, 0x28, 0x28)
        else:
            r = p.add_run(txt)
            r.font.name = 'Calibri'
            r.font.size = Pt(font_size)
            r.bold = el.get("bold", False)
            r.italic = el.get("italic", False)
            if el.get("underline", False):
                r.underline = True

        # Target cell — blank
        tp = tgt.paragraphs[0]
        tr = tp.add_run("")
        tr.font.name = 'Calibri'
        tr.font.size = Pt(font_size)

    # Build table
    add_col_header()
    total = 0

    for page in pages:
        page_num = page.get("pageNumber", 1)
        elements = page.get("elements", [])

        add_page_header(f"Page {page_num}")

        if not elements:
            fake_el = {"type": "paragraph", "text": "[No extractable content on this page]", "italic": True}
            add_content_row(fake_el, False, page_num)
        else:
            for idx, el in enumerate(elements):
                add_content_row(el, idx % 2 == 0, page_num)
                total += 1

    # Save to buffer
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf, total


# ═════════════════════════════════════════════════════════════
# UI
# ═════════════════════════════════════════════════════════════

# ── Step 1: API Key ───────────────────────────────────────────
st.markdown("### 🔑 Step 1 — Enter Your Claude API Key")
st.markdown("""
<div class="tip-box">
Your API key is used only for this session and is never stored.
Get yours free at <a href="https://console.anthropic.com" target="_blank">console.anthropic.com</a>
</div>
""", unsafe_allow_html=True)

api_key = st.text_input(
    "Claude API Key",
    type="password",
    placeholder="sk-ant-...",
    label_visibility="collapsed"
)

st.divider()

# ── Step 2: Upload files ──────────────────────────────────────
st.markdown("### 📁 Step 2 — Upload Files")
col1, col2 = st.columns(2)

with col1:
    st.markdown("**Source PDF**")
    pdf_file = st.file_uploader(
        "Upload PDF",
        type=["pdf"],
        label_visibility="collapsed",
        key="pdf_upload"
    )
    if pdf_file:
        st.success(f"✓ {pdf_file.name}")

with col2:
    st.markdown("**Word Template (.docx)**")
    tpl_file = st.file_uploader(
        "Upload Template",
        type=["docx"],
        label_visibility="collapsed",
        key="tpl_upload"
    )
    if tpl_file:
        st.success(f"✓ {tpl_file.name}")

st.divider()

# ── Step 3: Settings ──────────────────────────────────────────
st.markdown("### ⚙️ Step 3 — Settings")
col3, col4, col5 = st.columns(3)

with col3:
    source_lang = st.text_input("Source Language", value="English", placeholder="e.g. English")
with col4:
    target_lang = st.text_input("Target Language", value="", placeholder="e.g. French")
with col5:
    font_size = st.selectbox("Font Size", options=[10, 11, 12], index=0, format_func=lambda x: f"{x}pt")

st.markdown("""
<div class="tip-box">
⚠️ <strong>No transcription:</strong> Text is extracted directly from the PDF by Claude AI.
Non-extractable content (stamps, images, seals) will be flagged as
<code>[NON-EXTRACTABLE CONTENT — Page X]</code>.
</div>
""", unsafe_allow_html=True)

st.divider()

# ── Step 4: Generate ──────────────────────────────────────────
st.markdown("### 🚀 Step 4 — Generate")

ready = api_key and pdf_file and tpl_file

if not ready:
    missing = []
    if not api_key:   missing.append("API Key")
    if not pdf_file:  missing.append("PDF file")
    if not tpl_file:  missing.append("Word template")
    st.info(f"Still needed: {', '.join(missing)}")

generate_btn = st.button("⚡ Generate Bilingual Table", disabled=not ready)

if generate_btn and ready:
    try:
        # Extract
        with st.spinner("📤 Sending PDF to Claude AI for extraction..."):
            pdf_bytes = pdf_file.read()
            extracted = extract_with_claude(pdf_bytes, api_key)

        pages = extracted.get("pages", [])
        total_pages = len(pages)

        if total_pages == 0:
            st.error("Could not extract content from the PDF. Please check the file and try again.")
        else:
            st.success(f"✓ Extracted {total_pages} page(s) of content")

            with st.expander("📋 Extraction summary"):
                for p in pages:
                    count = len(p.get("elements", []))
                    st.write(f"• Page {p.get('pageNumber', '?')}: {count} elements")

            # Build DOCX
            with st.spinner("📝 Building bilingual Word document..."):
                doc_buf, total_rows = build_docx(pages, source_lang, target_lang or "Translation", font_size)

            # Done
            base_name = pdf_file.name.replace(".pdf", "").replace(".PDF", "")
            output_name = f"{base_name}_bilingual.docx"

            st.markdown(f"""
            <div class="success-box">
                <h3 style="color:#2e7d32;">✅ Bilingual Table Ready!</h3>
                <p style="color:#546e7a;">{total_pages} page section(s) &nbsp;·&nbsp; {total_rows} content rows &nbsp;·&nbsp; Calibri {font_size}pt &nbsp;·&nbsp; Target column ready for translation</p>
            </div>
            """, unsafe_allow_html=True)

            st.download_button(
                label="⬇️ Download Bilingual .docx",
                data=doc_buf,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    except anthropic.AuthenticationError:
        st.error("❌ Invalid API key. Please check your Claude API key and try again.")
    except anthropic.RateLimitError:
        st.error("❌ Rate limit reached. Please wait a moment and try again.")
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        st.exception(e)

# ── Footer ────────────────────────────────────────────────────
st.divider()
st.markdown("""
<p style="text-align:center;font-size:12px;color:#90a4ae;">
Tomedes OPEX DTP Team &nbsp;·&nbsp; Bilingual Table Generator &nbsp;·&nbsp; Powered by Claude AI
</p>
""", unsafe_allow_html=True)
