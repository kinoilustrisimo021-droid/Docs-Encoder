"""
Streamlit Template Filler for .docx forms
------------------------------------------
Upload one or more Word (.docx) templates that contain placeholders written
as {{PLACEHOLDER_NAME}} (e.g. {{Client_Name}}, {{BRAND}}, {{UNIT_DESCRIPTION}}).

The app will:
  1. Scan every uploaded template and auto-detect all unique {{...}} placeholders -
     paragraphs, tables/nested tables, headers/footers, AND text boxes/floating
     shapes are all scanned (Word often hides content in text boxes that most
     simple docx-templating scripts miss).
  2. Build a dynamic input form for those placeholders (fill once, applies to
     every uploaded template that shares the same field name).
  3. On submit, immediately generate the final .docx file(s) with ONLY the
     placeholder text replaced - everything else (fonts, bold, tables, spacing,
     letterhead, etc.) stays exactly as in the original template - and offer
     them for download right away. If more than one file was uploaded, they
     are also bundled into a single .zip for a one-click download.

Nothing is written to disk and no global/shared state is used, so this is
safe to run for many concurrent users on the same Streamlit server - all
processing happens in-memory, scoped to each user's own session.

UI NOTE: only the presentation layer below has changed (dark / glassmorphism
theme + animated background). All document-processing logic is untouched.
"""

import io
import re
import zipfile
from dataclasses import dataclass, field

import streamlit as st
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

# --------------------------------------------------------------------------
# Config
# --------------------------------------------------------------------------
st.set_page_config(page_title="Docx Template Filler", page_icon="📄", layout="centered")

PLACEHOLDER_PATTERN = re.compile(r"\{\{\s*([A-Za-z0-9_]+)\s*\}\}")

# Field-name hints -> nicer input widget / label. Purely cosmetic, matching is
# substring-based and case-insensitive, so it auto-applies to new templates too.
DATE_HINTS = ("DATE",)
LONG_TEXT_HINTS = ("ADDRESS", "DESCRIPTION", "REMARKS", "REASON")


# --------------------------------------------------------------------------
# Helpers: find every paragraph anywhere in the .docx - body text, tables
# (incl. nested tables), headers/footers, AND text boxes / floating shapes.
#
# Text boxes are the tricky part: Word stores their paragraphs nested deep
# inside <w:drawing>/<mc:AlternateContent> (and often duplicated once for
# modern DrawingML + once for a legacy VML fallback, both need filling).
# python-docx's own `.paragraphs` / `.tables` helpers only walk *direct*
# body children, so they silently skip anything inside a text box. Instead
# of a hand-rolled tree walk, we just ask lxml for every <w:p> element
# anywhere below the body/header/footer, no matter how deeply nested - this
# reaches table cells, nested tables, and text boxes uniformly.
# --------------------------------------------------------------------------
def _iter_paragraphs_in(xml_element, parent):
    for p_elm in xml_element.iter(qn("w:p")):
        yield Paragraph(p_elm, parent)


def iter_all_paragraphs(doc: Document):
    """Every paragraph in the whole document: body, tables, nested tables,
    text boxes/shapes, and every header/footer of every section.

    Important: python-docx silently creates a brand-new (blank) header/footer
    XML part the moment you so much as *read* `section.header` / `.footer`
    when the original template never defined one for that section (it was
    just inheriting the default blank header/footer). That bloats the saved
    file with parts that were never in the original, so we check
    `is_linked_to_previous` first (a safe, read-only check) and only touch
    a header/footer if it actually already exists in the template."""
    yield from _iter_paragraphs_in(doc.element.body, doc)

    for section in doc.sections:
        for hf in (
            section.header,
            section.footer,
            section.first_page_header,
            section.first_page_footer,
            section.even_page_header,
            section.even_page_footer,
        ):
            if hf is None or hf.is_linked_to_previous:
                continue
            yield from _iter_paragraphs_in(hf._element, hf)


def find_placeholders(doc: Document) -> list:
    """Return unique placeholder keys, in first-seen order."""
    seen = {}
    for p in iter_all_paragraphs(doc):
        text = "".join(r.text for r in p.runs)
        for m in PLACEHOLDER_PATTERN.finditer(text):
            seen.setdefault(m.group(1).strip(), None)
    return list(seen.keys())


def replace_placeholders_in_paragraph(paragraph: Paragraph, mapping: dict) -> None:
    """Replace {{KEY}} occurrences in a paragraph with mapping[KEY], even when
    Word has split the placeholder text across multiple runs. Formatting of
    surrounding / untouched text is preserved exactly; the replacement text
    inherits the formatting of the run in which the placeholder starts."""
    runs = paragraph.runs
    if not runs:
        return

    run_texts = [r.text for r in runs]
    full_text = "".join(run_texts)
    if "{{" not in full_text:
        return

    matches = list(PLACEHOLDER_PATTERN.finditer(full_text))
    if not matches:
        return

    # character-offset span owned by each run within full_text
    boundaries = []
    pos = 0
    for t in run_texts:
        boundaries.append((pos, pos + len(t)))
        pos += len(t)

    # which run "owns" each match (the run the match starts in) -> gets the
    # replacement text inserted; other runs touched by the match just lose
    # their portion of the placeholder text.
    owner_of = {}
    for m in matches:
        owner_idx = len(boundaries) - 1
        for idx, (s, e) in enumerate(boundaries):
            if s <= m.start() < e:
                owner_idx = idx
                break
        owner_of[m] = owner_idx

    new_texts = []
    for idx, (s, e) in enumerate(boundaries):
        overlapping = [m for m in matches if m.start() < e and m.end() > s]
        overlapping.sort(key=lambda m: m.start())
        out_parts = []
        cursor = s
        for m in overlapping:
            if m.start() > cursor:
                out_parts.append(full_text[cursor:min(m.start(), e)])
            if owner_of[m] == idx:
                key = m.group(1).strip()
                out_parts.append(str(mapping.get(key, m.group(0))))
            cursor = min(m.end(), e)
        if cursor < e:
            out_parts.append(full_text[cursor:e])
        new_texts.append("".join(out_parts))

    for run, new_text in zip(runs, new_texts):
        if run.text != new_text:
            run.text = new_text


def fill_template(file_bytes: bytes, mapping: dict) -> bytes:
    """Return a new .docx (as bytes) with every {{placeholder}} replaced."""
    doc = Document(io.BytesIO(file_bytes))
    for p in iter_all_paragraphs(doc):
        replace_placeholders_in_paragraph(p, mapping)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def label_for(key: str) -> str:
    return key.replace("_", " ").strip().title()


def widget_for(key: str, current_value: str):
    upper = key.upper()
    if any(h in upper for h in DATE_HINTS):
        # Date fields use a calendar picker and are stored in the requested
        # MM/DD/YYYY format before being inserted into the document.
        current_date = None
        if current_value:
            try:
                from datetime import datetime
                current_date = datetime.strptime(current_value, "%m/%d/%Y").date()
            except ValueError:
                current_date = None

        selected_date = st.date_input(
            label_for(key),
            value=current_date,
            format="MM/DD/YYYY",
            key=f"field_{key}",
        )
        if selected_date is None:
            return ""
        return selected_date.strftime("%m/%d/%Y")
    if any(h in upper for h in LONG_TEXT_HINTS):
        return st.text_area(label_for(key), value=current_value, key=f"field_{key}", height=80)
    return st.text_input(label_for(key), value=current_value, key=f"field_{key}")


# --------------------------------------------------------------------------
# Session state
# --------------------------------------------------------------------------
@dataclass
class AppState:
    step: str = "upload"          # upload -> fill -> done
    templates: list = field(default_factory=list)   # [{"name": str, "bytes": bytes}]
    placeholders: list = field(default_factory=list)
    values: dict = field(default_factory=dict)
    filled: list = field(default_factory=list)      # [{"name": str, "bytes": bytes}]


if "state" not in st.session_state:
    st.session_state.state = AppState()

state: AppState = st.session_state.state


def go_to(step: str):
    state.step = step


def reset_all():
    st.session_state.state = AppState()


# ==========================================================================
# THEME — dark / glassmorphism / animated background
# ==========================================================================
def inject_theme():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700;800&family=Inter:wght@400;500;600&display=swap');

        :root{
            --bg-0:#05060c;
            --bg-1:#080a16;
            --panel-bg: rgba(17, 19, 34, 0.55);
            --panel-border: rgba(255,255,255,0.09);
            --text-primary:#eef0fb;
            --text-muted: rgba(222,225,245,0.55);
            --glow-purple: rgba(168, 85, 247, 0.45);
            --glow-blue: rgba(59, 130, 246, 0.40);
            --glow-cyan: rgba(34, 211, 238, 0.40);
            --accent-grad: linear-gradient(120deg, #7c3aed 0%, #3b82f6 55%, #22d3ee 100%);
        }

        /* Streamlit paints its own built-in widgets (text inputs, text
           areas, date inputs, etc.) using ITS OWN theme CSS variables, not
           the ones defined above. Whichever theme got resolved for a given
           visitor (their own cached Light/Dark choice, "Auto", or this
           app's Custom theme), those widgets read --primary-color,
           --background-color, --secondary-background-color and --text-color
           to decide what to paint. Overriding those variables directly -
           instead of guessing at internal wrapper class names, which can
           change between Streamlit versions - forces every built-in widget
           to use this app's colors regardless of which theme Streamlit
           thinks is active for that particular visitor. */
        :root, .stApp, [data-testid="stAppViewContainer"]{
            --primary-color: #7c3aed !important;
            --background-color: #05060c !important;
            --secondary-background-color: rgba(255,255,255,0.03) !important;
            --text-color: #eef0fb !important;
        }

        html, body, [class*="css"]{
            font-family: 'Inter', sans-serif;
        }

        /* ---------- base app background ---------- */
        .stApp{
            background:
                radial-gradient(circle at 15% 20%, rgba(124,58,237,0.16), transparent 40%),
                radial-gradient(circle at 85% 15%, rgba(34,211,238,0.12), transparent 42%),
                radial-gradient(circle at 50% 90%, rgba(59,130,246,0.14), transparent 45%),
                linear-gradient(180deg, var(--bg-0) 0%, var(--bg-1) 100%);
            background-attachment: fixed;
            overflow-x: hidden;
        }

        /* hide default streamlit chrome for a minimal shell */
        #MainMenu{visibility:hidden;}
        footer{visibility:hidden;}
        header[data-testid="stHeader"]{
            background: transparent;
            box-shadow: none;
        }
        [data-testid="stToolbar"]{visibility:hidden;}
        [data-testid="stDecoration"]{display:none;}

        /* ---------- vertical + horizontal centering ---------- */
        [data-testid="stAppViewContainer"] > .main{
            display:flex;
            align-items:center;
            justify-content:center;
            min-height: 100vh;
            padding: 2rem 1rem;
        }

        /* ---------- the glass card that holds everything ---------- */
        .main .block-container{
            max-width: 620px;
            width: 100%;
            background: var(--panel-bg);
            backdrop-filter: blur(22px);
            -webkit-backdrop-filter: blur(22px);
            border: 1px solid var(--panel-border);
            border-radius: 28px;
            padding: 3rem 2.6rem 2.6rem 2.6rem;
            position: relative;
            z-index: 2;
            box-shadow:
                0 0 0 1px rgba(255,255,255,0.02) inset,
                0 30px 90px -20px rgba(0,0,0,0.65),
                0 0 120px -40px var(--glow-purple);
            animation: card-glow 8s ease-in-out infinite;
        }

        @keyframes card-glow{
            0%, 100% { box-shadow: 0 0 0 1px rgba(255,255,255,0.02) inset, 0 30px 90px -20px rgba(0,0,0,0.65), 0 0 100px -45px var(--glow-purple); }
            50% { box-shadow: 0 0 0 1px rgba(255,255,255,0.02) inset, 0 30px 90px -20px rgba(0,0,0,0.65), 0 0 130px -35px var(--glow-cyan); }
        }

        /* ---------- background decorative layer ---------- */
        .bg-decor{
            position: fixed;
            inset: 0;
            width: 100vw;
            height: 100vh;
            overflow: hidden;
            z-index: 0;
            pointer-events: none;
        }
        .bg-orb{
            position:absolute;
            border-radius:50%;
            filter: blur(6px);
            opacity:0.55;
        }
        .orb-a{ width:320px; height:320px; left:6%; top:8%;
            background: radial-gradient(circle, var(--glow-purple), transparent 70%);
            animation: drift-a 26s ease-in-out infinite; }
        .orb-b{ width:260px; height:260px; right:8%; top:18%;
            background: radial-gradient(circle, var(--glow-cyan), transparent 70%);
            animation: drift-b 30s ease-in-out infinite; }
        .orb-c{ width:380px; height:380px; left:20%; bottom:-8%;
            background: radial-gradient(circle, var(--glow-blue), transparent 70%);
            animation: drift-c 34s ease-in-out infinite; }
        .orb-d{ width:200px; height:200px; right:14%; bottom:6%;
            background: radial-gradient(circle, var(--glow-purple), transparent 70%);
            animation: drift-a 22s ease-in-out infinite reverse; }

        @keyframes drift-a{
            0%,100%{ transform: translate(0,0) scale(1); }
            50%{ transform: translate(40px,50px) scale(1.08); }
        }
        @keyframes drift-b{
            0%,100%{ transform: translate(0,0) scale(1); }
            50%{ transform: translate(-50px,35px) scale(0.94); }
        }
        @keyframes drift-c{
            0%,100%{ transform: translate(0,0) scale(1); }
            50%{ transform: translate(35px,-45px) scale(1.06); }
        }

        .bg-particle{
            position:absolute;
            width:4px; height:4px;
            border-radius:50%;
            background: rgba(210,220,255,0.7);
            box-shadow: 0 0 8px 2px rgba(150,180,255,0.5);
        }
        .p1{ left:12%; top:30%; animation: float-p1 14s ease-in-out infinite; }
        .p2{ left:78%; top:22%; animation: float-p2 17s ease-in-out infinite; }
        .p3{ left:35%; top:70%; animation: float-p3 12s ease-in-out infinite; }
        .p4{ left:88%; top:60%; animation: float-p1 19s ease-in-out infinite reverse; }
        .p5{ left:55%; top:12%; animation: float-p2 15s ease-in-out infinite; }
        .p6{ left:22%; top:85%; animation: float-p3 20s ease-in-out infinite reverse; }

        @keyframes float-p1{
            0%,100%{ transform: translate(0,0); opacity:0.3; }
            50%{ transform: translate(18px,-26px); opacity:0.9; }
        }
        @keyframes float-p2{
            0%,100%{ transform: translate(0,0); opacity:0.25; }
            50%{ transform: translate(-22px,20px); opacity:0.8; }
        }
        @keyframes float-p3{
            0%,100%{ transform: translate(0,0); opacity:0.35; }
            50%{ transform: translate(14px,24px); opacity:0.85; }
        }

        .bg-ring{
            position:absolute;
            border-radius:50%;
            border: 1px solid rgba(150,170,255,0.14);
        }
        .r1{ width:520px; height:520px; left:-120px; top:-140px; animation: spin-cw 70s linear infinite; }
        .r2{ width:640px; height:640px; right:-200px; bottom:-220px; border-color: rgba(120,230,255,0.10); animation: spin-ccw 90s linear infinite; }

        @keyframes spin-cw{ from{ transform: rotate(0deg);} to{ transform: rotate(360deg);} }
        @keyframes spin-ccw{ from{ transform: rotate(0deg);} to{ transform: rotate(-360deg);} }

        .bg-doc{
            position:absolute;
            font-size: 1.6rem;
            opacity:0.18;
            filter: drop-shadow(0 0 6px rgba(140,160,255,0.4));
        }
        .d1{ left:10%; top:55%; animation: float-doc 9s ease-in-out infinite; }
        .d2{ right:16%; top:38%; animation: float-doc 11s ease-in-out infinite reverse; }
        .d3{ left:48%; top:82%; animation: float-doc 8s ease-in-out infinite; }

        @keyframes float-doc{
            0%,100%{ transform: translateY(0) rotate(-4deg); }
            50%{ transform: translateY(-16px) rotate(4deg); }
        }

        /* ---------- main icon + title ---------- */
        .app-icon{
            text-align:center;
            font-size: 3.4rem;
            line-height:1;
            margin-bottom: 0.6rem;
            filter: drop-shadow(0 0 18px var(--glow-purple));
            animation: float-doc 6s ease-in-out infinite;
        }
        .app-title{
            text-align:center;
            font-family:'Sora', sans-serif;
            font-weight:700;
            font-size: 1.9rem;
            letter-spacing:0.2px;
            margin: 0 0 2.2rem 0;
            background: var(--accent-grad);
            -webkit-background-clip: text;
            background-clip: text;
            color: transparent;
        }
        .step-title{
            text-align:center;
            font-family:'Sora', sans-serif;
            font-weight:600;
            font-size: 1.3rem;
            color: var(--text-primary);
            margin: 0 0 1.6rem 0;
        }
        .ghost-back{
            text-align:center;
            margin-bottom: 0.6rem;
        }

        /* ---------- file uploader as a glowing dropzone ---------- */
        [data-testid="stFileUploaderDropzone"]{
            background: rgba(255,255,255,0.02);
            border: 1.5px dashed rgba(150,170,255,0.35);
            border-radius: 20px;
            padding: 1.2rem;
            transition: border-color 0.3s ease, box-shadow 0.3s ease;
            animation: pulse-drop 4.5s ease-in-out infinite;
        }
        [data-testid="stFileUploaderDropzone"]:hover{
            border-color: rgba(180,200,255,0.6);
        }
        @keyframes pulse-drop{
            0%,100%{ box-shadow: 0 0 0px 0px rgba(124,58,237,0.0); }
            50%{ box-shadow: 0 0 26px 4px rgba(124,58,237,0.22); }
        }
        [data-testid="stFileUploaderDropzoneInstructions"] div,
        [data-testid="stFileUploaderDropzoneInstructions"] span{
            color: var(--text-muted) !important;
            font-size: 0.82rem !important;
        }
        [data-testid="stFileUploaderDropzoneInstructions"] svg{
            fill: rgba(180,200,255,0.7) !important;
        }
        /* relabel the native "Browse files" button to "Upload" */
        [data-testid="stFileUploaderDropzone"] button{
            background: var(--accent-grad) !important;
            color: #fff !important;
            border: none !important;
            border-radius: 12px !important;
            padding: 0.5rem 1.3rem !important;
            font-weight: 600 !important;
            box-shadow: 0 6px 20px -6px rgba(124,58,237,0.6);
        }
        [data-testid="stFileUploaderDropzone"] button p{
            visibility:hidden;
            position:relative;
        }
        [data-testid="stFileUploaderDropzone"] button p::after{
            content: "Upload";
            visibility: visible;
            position:absolute;
            left:0; top:0; right:0;
        }
        [data-testid="stFileUploaderFile"]{
            background: rgba(255,255,255,0.04);
            border: 1px solid var(--panel-border);
            border-radius: 12px;
        }

        /* ---------- generic buttons ---------- */
        .stButton button, .stDownloadButton button, .stFormSubmitButton button{
            border-radius: 12px !important;
            font-weight: 600 !important;
            border: 1px solid var(--panel-border) !important;
            transition: transform 0.15s ease, box-shadow 0.15s ease;
        }
        .stButton button:hover, .stDownloadButton button:hover, .stFormSubmitButton button:hover{
            transform: translateY(-1px);
        }
        button[kind="primary"], button[kind="primaryFormSubmit"]{
            background: var(--accent-grad) !important;
            border: none !important;
            box-shadow: 0 8px 24px -8px rgba(124,58,237,0.6);
        }
        button[kind="secondary"], button[kind="secondaryFormSubmit"]{
            background: rgba(255,255,255,0.04) !important;
            color: var(--text-primary) !important;
        }

        /* ---------- inputs ----------
           Streamlit lets each visitor's OWN browser remember a Light/Dark
           theme choice in local storage, and that choice silently overrides
           this app's config.toml theme for that visitor - even though the
           menu that sets it is hidden below. Rather than depend on that
           override never happening, every layer of every input widget
           (the outer wrapper, any inner BaseWeb container, and the actual
           input/textarea element) is force-styled here with !important, so
           the app looks identical no matter what theme Streamlit resolves
           to for a given visitor. */
        [data-testid="stTextInput"] [data-baseweb],
        [data-testid="stTextArea"] [data-baseweb],
        [data-testid="stDateInput"] [data-baseweb],
        .stTextInput input, .stTextArea textarea, .stDateInput input{
            background: rgba(255,255,255,0.03) !important;
            background-color: rgba(255,255,255,0.03) !important;
            border: 1px solid var(--panel-border) !important;
            border-radius: 10px !important;
            color: var(--text-primary) !important;
            box-shadow: none !important;
        }
        .stTextInput input::placeholder, .stTextArea textarea::placeholder{
            color: var(--text-muted) !important;
            opacity: 1 !important;
        }
        [data-testid="stTextInput"]:focus-within [data-baseweb],
        [data-testid="stTextArea"]:focus-within [data-baseweb],
        [data-testid="stDateInput"]:focus-within [data-baseweb]{
            border-color: rgba(124,58,237,0.65) !important;
            box-shadow: 0 0 0 3px rgba(124,58,237,0.18) !important;
        }
        .stTextInput label, .stTextArea label, .stDateInput label,
        [data-testid="stWidgetLabel"] p{
            color: var(--text-muted) !important;
            font-size: 0.85rem !important;
        }

        /* the date-picker calendar renders in a portal outside the card
           (attached near <body>), so it needs its own explicit dark
           styling - it will not inherit anything scoped to .main above */
        div[data-baseweb="popover"]{
            background: var(--bg-1) !important;
            border: 1px solid var(--panel-border) !important;
            border-radius: 12px !important;
        }
        div[data-baseweb="popover"] *{
            color: var(--text-primary) !important;
        }
        div[data-baseweb="popover"] [aria-selected="true"]{
            background: var(--accent-grad) !important;
            color: #fff !important;
        }

        /* ---------- alerts, expanders ---------- */
        [data-testid="stAlert"]{
            background: rgba(255,255,255,0.04) !important;
            border: 1px solid var(--panel-border) !important;
            border-radius: 12px !important;
            color: var(--text-primary) !important;
        }
        [data-testid="stExpander"]{
            background: rgba(255,255,255,0.03) !important;
            border: 1px solid var(--panel-border) !important;
            border-radius: 12px !important;
        }
        [data-testid="stCaptionContainer"]{
            color: var(--text-muted) !important;
        }
        hr{ border-color: var(--panel-border) !important; }

        /* ---------- responsive ---------- */
        @media (max-width: 640px){
            .main .block-container{
                padding: 2.2rem 1.4rem;
                border-radius: 22px;
            }
            .app-icon{ font-size: 2.6rem; }
            .app-title{ font-size: 1.5rem; }
        }
        </style>

        <div class="bg-decor">
            <div class="bg-orb orb-a"></div>
            <div class="bg-orb orb-b"></div>
            <div class="bg-orb orb-c"></div>
            <div class="bg-orb orb-d"></div>
            <div class="bg-ring r1"></div>
            <div class="bg-ring r2"></div>
            <div class="bg-particle p1"></div>
            <div class="bg-particle p2"></div>
            <div class="bg-particle p3"></div>
            <div class="bg-particle p4"></div>
            <div class="bg-particle p5"></div>
            <div class="bg-particle p6"></div>
            <div class="bg-doc d1">📄</div>
            <div class="bg-doc d2">📄</div>
            <div class="bg-doc d3">📄</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


inject_theme()

# --------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------

# --- Step 1: Upload -------------------------------------------------------
if state.step == "upload":
    st.markdown('<div class="app-icon">📄</div>', unsafe_allow_html=True)
    st.markdown('<div class="app-title">Docx Template Filler</div>', unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "Upload template",
        type=["docx"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    if uploaded:
        templates = [{"name": f.name, "bytes": f.getvalue()} for f in uploaded]

        placeholders = {}
        bad_files = []
        for t in templates:
            try:
                doc = Document(io.BytesIO(t["bytes"]))
            except Exception:
                bad_files.append(t["name"])
                continue
            for key in find_placeholders(doc):
                placeholders.setdefault(key, None)

        if bad_files:
            st.error(f"Could not read: {', '.join(bad_files)}. Please re-upload a valid .docx file.")

        if placeholders:
            if st.button("Continue →", type="primary", use_container_width=True):
                state.templates = templates
                state.placeholders = list(placeholders.keys())
                state.values = {k: "" for k in state.placeholders}
                go_to("fill")
                st.rerun()
        else:
            st.warning("No `{{placeholder}}` fields were found in this file.")

# --- Step 2: Dynamic form -> generate & download immediately ---------------
elif state.step == "fill":
    st.markdown('<div class="step-title">Fill in the fields</div>', unsafe_allow_html=True)

    with st.form("fill_form"):
        new_values = {}
        for key in state.placeholders:
            new_values[key] = widget_for(key, state.values.get(key, ""))

        col1, col2 = st.columns([1, 1])
        with col1:
            back = st.form_submit_button("← Back", use_container_width=True)
        with col2:
            submitted = st.form_submit_button("Generate →", type="primary", use_container_width=True)

    if back:
        go_to("upload")
        st.rerun()

    if submitted:
        state.values = new_values
        filled = []
        for t in state.templates:
            try:
                out_bytes = fill_template(t["bytes"], state.values)
            except Exception as e:
                st.error(f"Failed to generate {t['name']}: {e}")
                continue
            stem = t["name"].rsplit(".", 1)[0]
            filled.append({"name": f"{stem}_filled.docx", "bytes": out_bytes})

        if filled:
            state.filled = filled
            go_to("done")
            st.rerun()

# --- Step 3: Download --------------------------------------------------------
elif state.step == "done":
    st.markdown('<div class="step-title">Ready to download</div>', unsafe_allow_html=True)

    filled = state.filled

    if len(filled) == 1:
        f = filled[0]
        st.download_button(
            "⬇️ Download " + f["name"],
            data=f["bytes"],
            file_name=f["name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )
    else:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for f in filled:
                zf.writestr(f["name"], f["bytes"])

        st.download_button(
            "⬇️ Download all as .zip",
            data=zip_buf.getvalue(),
            file_name="filled_documents.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True,
        )

        with st.expander("Download individually"):
            for f in filled:
                st.download_button(
                    f["name"],
                    data=f["bytes"],
                    file_name=f["name"],
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"dl_{f['name']}",
                    use_container_width=True,
                )

    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("✏️ Edit fields", use_container_width=True):
            go_to("fill")
            st.rerun()
    with col2:
        if st.button("🔄 Start over", use_container_width=True):
            reset_all()
            st.rerun()
