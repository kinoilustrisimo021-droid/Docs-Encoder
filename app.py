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
# Preserve the app's 200 MB upload limit.
st.set_option("server.maxUploadSize", 200)

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


# --------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------
st.markdown(
    """
    <style>
    /* ---------- Global canvas ---------- */
    html, body, [data-testid="stAppViewContainer"] {
        background:
            radial-gradient(circle at 50% 45%, rgba(61, 45, 130, 0.16), transparent 32%),
            radial-gradient(circle at 12% 18%, rgba(25, 109, 180, 0.10), transparent 26%),
            radial-gradient(circle at 88% 78%, rgba(0, 190, 210, 0.07), transparent 28%),
            #050814;
    }

    [data-testid="stAppViewContainer"] {
        overflow: hidden;
        position: relative;
    }

    [data-testid="stHeader"] {
        background: transparent;
    }

    [data-testid="stToolbar"],
    [data-testid="stDecoration"],
    footer {
        display: none !important;
    }

    /* ---------- Animated atmosphere ---------- */
    [data-testid="stAppViewContainer"]::before,
    [data-testid="stAppViewContainer"]::after {
        content: "";
        position: fixed;
        inset: -20%;
        pointer-events: none;
        z-index: 0;
        background-repeat: no-repeat;
    }

    [data-testid="stAppViewContainer"]::before {
        background-image:
            radial-gradient(circle at 18% 28%, rgba(103, 76, 255, .20) 0 2px, transparent 3px),
            radial-gradient(circle at 74% 22%, rgba(72, 191, 255, .16) 0 2px, transparent 3px),
            radial-gradient(circle at 83% 70%, rgba(57, 239, 255, .14) 0 2px, transparent 3px),
            radial-gradient(circle at 31% 78%, rgba(151, 90, 255, .12) 0 1.5px, transparent 3px);
        background-size: 310px 260px, 370px 310px, 430px 360px, 280px 300px;
        animation: particleDrift 24s linear infinite;
        opacity: .8;
    }

    [data-testid="stAppViewContainer"]::after {
        background-image:
            radial-gradient(ellipse 430px 180px at 50% 50%,
                transparent 56%,
                rgba(94, 83, 255, .09) 56.3%,
                transparent 56.8%),
            radial-gradient(ellipse 620px 270px at 50% 50%,
                transparent 67%,
                rgba(41, 196, 255, .065) 67.2%,
                transparent 67.7%);
        animation: orbitSpin 32s linear infinite;
        opacity: .9;
    }

    .ambient {
        position: fixed;
        border-radius: 50%;
        pointer-events: none;
        z-index: 0;
        filter: blur(2px);
    }

    .orb-a {
        width: 170px;
        height: 170px;
        left: 7vw;
        top: 13vh;
        background: radial-gradient(circle, rgba(117, 76, 255, .18), transparent 68%);
        animation: orbFloatA 18s ease-in-out infinite;
    }

    .orb-b {
        width: 220px;
        height: 220px;
        right: 5vw;
        bottom: 8vh;
        background: radial-gradient(circle, rgba(38, 196, 255, .13), transparent 70%);
        animation: orbFloatB 22s ease-in-out infinite;
    }

    .orb-c {
        width: 100px;
        height: 100px;
        right: 20vw;
        top: 11vh;
        background: radial-gradient(circle, rgba(173, 91, 255, .11), transparent 68%);
        animation: orbFloatC 15s ease-in-out infinite;
    }

    .doc-float {
        position: fixed;
        color: rgba(124, 175, 255, .10);
        font-size: 42px;
        line-height: 1;
        pointer-events: none;
        z-index: 0;
        filter: drop-shadow(0 0 14px rgba(74, 181, 255, .10));
    }

    .doc-one {
        left: 13vw;
        bottom: 18vh;
        animation: docFloatOne 16s ease-in-out infinite;
    }

    .doc-two {
        right: 14vw;
        top: 22vh;
        font-size: 34px;
        animation: docFloatTwo 19s ease-in-out infinite;
    }

    .doc-three {
        left: 25vw;
        top: 13vh;
        font-size: 24px;
        opacity: .65;
        animation: docFloatThree 13s ease-in-out infinite;
    }

    @keyframes particleDrift {
        0%   { transform: translate3d(0, 0, 0); }
        50%  { transform: translate3d(22px, -28px, 0); }
        100% { transform: translate3d(0, 0, 0); }
    }

    @keyframes orbitSpin {
        from { transform: rotate(0deg) scale(1); }
        50%  { transform: rotate(180deg) scale(1.025); }
        to   { transform: rotate(360deg) scale(1); }
    }

    @keyframes orbFloatA {
        0%, 100% { transform: translate3d(0, 0, 0); }
        50% { transform: translate3d(55px, 35px, 0); }
    }

    @keyframes orbFloatB {
        0%, 100% { transform: translate3d(0, 0, 0); }
        50% { transform: translate3d(-48px, -30px, 0); }
    }

    @keyframes orbFloatC {
        0%, 100% { transform: translate3d(0, 0, 0); }
        50% { transform: translate3d(-28px, 45px, 0); }
    }

    @keyframes docFloatOne {
        0%, 100% { transform: translate3d(0, 0, 0) rotate(-7deg); }
        50% { transform: translate3d(22px, -32px, 0) rotate(4deg); }
    }

    @keyframes docFloatTwo {
        0%, 100% { transform: translate3d(0, 0, 0) rotate(8deg); }
        50% { transform: translate3d(-25px, 28px, 0) rotate(-4deg); }
    }

    @keyframes docFloatThree {
        0%, 100% { transform: translate3d(0, 0, 0) rotate(0deg); }
        50% { transform: translate3d(18px, 18px, 0) rotate(10deg); }
    }

    @keyframes uploaderPulse {
        0%, 100% {
            box-shadow:
                0 0 0 1px rgba(122, 93, 255, .16),
                0 0 28px rgba(75, 117, 255, .06),
                inset 0 0 28px rgba(102, 74, 255, .025);
        }
        50% {
            box-shadow:
                0 0 0 1px rgba(101, 184, 255, .25),
                0 0 42px rgba(73, 126, 255, .12),
                inset 0 0 36px rgba(99, 75, 255, .045);
        }
    }

    /* ---------- Main layout ---------- */
    .block-container {
        position: relative;
        z-index: 2;
        max-width: 760px !important;
        min-height: 100vh;
        padding-top: 0 !important;
        padding-bottom: 0 !important;
        display: flex;
        align-items: center;
    }

    [data-testid="stVerticalBlock"] {
        width: 100%;
    }

    .hero {
        width: min(680px, 92vw);
        margin: auto;
        padding: 44px 46px 48px;
        text-align: center;
        border: 1px solid rgba(143, 158, 210, .14);
        border-radius: 30px;
        background:
            linear-gradient(145deg, rgba(17, 22, 40, .72), rgba(7, 12, 26, .58));
        box-shadow:
            0 32px 90px rgba(0, 0, 0, .45),
            0 0 70px rgba(83, 74, 190, .08),
            inset 0 1px 0 rgba(255,255,255,.035);
        backdrop-filter: blur(24px);
        -webkit-backdrop-filter: blur(24px);
    }

    .doc-icon {
        width: 78px;
        height: 78px;
        margin: 0 auto 22px;
        display: grid;
        place-items: center;
        border-radius: 22px;
        background: linear-gradient(145deg, rgba(112, 89, 255, .18), rgba(45, 191, 255, .10));
        border: 1px solid rgba(140, 153, 255, .22);
        box-shadow:
            0 0 38px rgba(100, 84, 255, .13),
            inset 0 0 20px rgba(63, 199, 255, .035);
        color: #dce8ff;
        font-size: 42px;
        animation: docIconFloat 5.5s ease-in-out infinite;
    }

    @keyframes docIconFloat {
        0%, 100% { transform: translateY(0); }
        50% { transform: translateY(-5px); }
    }

    .hero h1 {
        margin: 0 !important;
        color: #f4f7ff;
        font-size: clamp(2rem, 5vw, 3.15rem);
        line-height: 1.05;
        font-weight: 700;
        letter-spacing: -0.045em;
        text-shadow: 0 0 30px rgba(130, 147, 255, .13);
    }

    /* ---------- Streamlit uploader ---------- */
    .uploader-shell {
        margin: 34px auto 0;
        animation: uploaderPulse 5s ease-in-out infinite;
        border-radius: 22px;
    }

    [data-testid="stFileUploader"] {
        margin: 0 !important;
    }

    [data-testid="stFileUploader"] section {
        border: 1px dashed rgba(117, 157, 226, .34) !important;
        border-radius: 22px !important;
        background:
            linear-gradient(145deg, rgba(16, 24, 44, .76), rgba(9, 15, 30, .70)) !important;
        min-height: 190px;
        padding: 32px 24px !important;
        transition: border-color .25s ease, background .25s ease, transform .25s ease;
    }

    [data-testid="stFileUploader"] section:hover {
        border-color: rgba(112, 190, 255, .56) !important;
        background: linear-gradient(145deg, rgba(19, 29, 52, .84), rgba(10, 18, 35, .76)) !important;
        transform: translateY(-1px);
    }

    [data-testid="stFileUploader"] section > div {
        gap: 8px;
    }

    [data-testid="stFileUploader"] svg {
        color: #9caeff !important;
        filter: drop-shadow(0 0 10px rgba(102, 126, 255, .35));
    }

    [data-testid="stFileUploader"] button {
        border: 1px solid rgba(125, 151, 255, .34) !important;
        background: linear-gradient(135deg, rgba(101, 78, 255, .23), rgba(31, 157, 220, .16)) !important;
        color: #edf4ff !important;
        border-radius: 12px !important;
        font-weight: 600 !important;
        transition: all .2s ease;
        box-shadow: 0 8px 22px rgba(38, 87, 180, .10);
    }

    [data-testid="stFileUploader"] button:hover {
        border-color: rgba(119, 210, 255, .60) !important;
        box-shadow: 0 0 24px rgba(77, 153, 255, .16);
    }

    [data-testid="stFileUploader"] small,
    [data-testid="stFileUploader"] label {
        color: rgba(218, 227, 247, .72) !important;
    }

    [data-testid="stFileUploader"] [data-testid="stMarkdownContainer"] p {
        color: rgba(225, 234, 252, .72) !important;
    }

    /* Hide the default uploader label while retaining the accessible widget. */
    [data-testid="stFileUploader"] > label {
        display: none !important;
    }

    /* Keep downstream form/download UI functional but visually consistent. */
    .stTextInput input,
    .stTextArea textarea,
    [data-testid="stDateInput"] input {
        background: rgba(10, 16, 31, .72) !important;
        border: 1px solid rgba(126, 145, 193, .18) !important;
        color: #eef3ff !important;
        border-radius: 12px !important;
    }

    .stButton button,
    .stDownloadButton button {
        border-radius: 12px !important;
    }

    @media (max-width: 640px) {
        .block-container {
            padding-left: 16px !important;
            padding-right: 16px !important;
        }

        .hero {
            padding: 32px 18px 34px;
            border-radius: 24px;
        }

        [data-testid="stFileUploader"] section {
            min-height: 170px;
            padding: 24px 14px !important;
        }

        .doc-one { left: 4vw; }
        .doc-two { right: 5vw; }
        .doc-three { left: 12vw; }
    }

    @media (prefers-reduced-motion: reduce) {
        .ambient,
        .doc-float,
        .doc-icon,
        [data-testid="stAppViewContainer"]::before,
        [data-testid="stAppViewContainer"]::after,
        .uploader-shell {
            animation: none !important;
        }
    }
    </style>

    <div class="ambient orb-a"></div>
    <div class="ambient orb-b"></div>
    <div class="ambient orb-c"></div>
    <div class="doc-float doc-one">▱</div>
    <div class="doc-float doc-two">▱</div>
    <div class="doc-float doc-three">▱</div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero">
        <div class="doc-icon" aria-hidden="true">📄</div>
        <h1>Docx Template Filler</h1>
        <div class="uploader-shell">
    """,
    unsafe_allow_html=True,
)

# Single DOCX upload surface. The existing document-processing functions and
# state flow below remain unchanged.
uploaded = st.file_uploader(
    "Upload",
    type=["docx"],
    accept_multiple_files=False,
    help=None,
)

st.markdown(
    """
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

if uploaded:
    templates = [{"name": uploaded.name, "bytes": uploaded.getvalue()}]

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
        if st.button("Continue to form →", type="primary", use_container_width=True):
            state.templates = templates
            state.placeholders = list(placeholders.keys())
            state.values = {k: "" for k in state.placeholders}
            go_to("fill")
            st.rerun()
    else:
        st.warning(
            "No `{{placeholder}}` fields were found in the uploaded file(s). "
            "Make sure the template uses double curly braces, e.g. {{Client_Name}}."
        )

# --- Step 2: Dynamic form -> generate & download immediately ---------------
elif state.step == "fill":
    st.subheader("Fill in the fields")
    st.caption(
        f"{len(state.templates)} template(s) loaded · {len(state.placeholders)} field(s) · "
        "any field left blank will just be left blank in the final document."
    )

    with st.form("fill_form"):
        new_values = {}
        for key in state.placeholders:
            new_values[key] = widget_for(key, state.values.get(key, ""))

        col1, col2 = st.columns([1, 1])
        with col1:
            back = st.form_submit_button("← Back")
        with col2:
            submitted = st.form_submit_button("✅ Generate & download →", type="primary")

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
    st.subheader("✅ Done! Your document(s) are ready")

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

        st.divider()
        st.caption("Or download individually:")
        for f in filled:
            st.download_button(
                f["name"],
                data=f["bytes"],
                file_name=f["name"],
                key=f"dl_{f['name']}",
            )

    st.divider()
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("✏️ Edit fields & regenerate", use_container_width=True):
            go_to("fill")
            st.rerun()
    with col2:
        if st.button("🔄 Start over with new file(s)", use_container_width=True):
            reset_all()
            st.rerun()
