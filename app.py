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
        return st.text_input(
            f"{label_for(key)} *", value=current_value, placeholder="e.g. January 5, 2026", key=f"field_{key}"
        )
    if any(h in upper for h in LONG_TEXT_HINTS):
        return st.text_area(f"{label_for(key)} *", value=current_value, key=f"field_{key}", height=80)
    return st.text_input(f"{label_for(key)} *", value=current_value, key=f"field_{key}")


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
st.title("📄 Docx Template Filler")
st.caption(
    "Upload your Word template(s) with {{placeholders}} → fill in the required fields → "
    "download the completed document(s) instantly, formatting untouched."
)

with st.sidebar:
    st.subheader("How it works")
    st.markdown(
        "1. **Upload** one or more `.docx` templates\n"
        "2. Fill out the **auto-generated form**\n"
        "3. **Download** the finished document(s) right away"
    )
    if state.step != "upload":
        st.divider()
        if st.button("🔄 Start over", use_container_width=True):
            reset_all()
            st.rerun()

# --- Step 1: Upload -------------------------------------------------------
if state.step == "upload":
    uploaded = st.file_uploader(
        "Upload template(s) (.docx)", type=["docx"], accept_multiple_files=True
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
            st.success(f"Found {len(placeholders)} field(s) across {len(templates)} file(s).")
            with st.expander("Detected fields", expanded=False):
                st.write(", ".join(f"{{{{{k}}}}}" for k in placeholders))

            if st.button("Continue to form →", type="primary"):
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
    st.subheader("Fill in the required fields")
    st.caption(f"{len(state.templates)} template(s) loaded · {len(state.placeholders)} field(s) to fill")

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
        missing = [k for k, v in new_values.items() if not str(v).strip()]
        state.values = new_values
        if missing:
            st.error(
                "Please fill in all required fields: "
                + ", ".join(label_for(k) for k in missing)
            )
        else:
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
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
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
