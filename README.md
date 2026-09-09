# Docx Template Filler (Streamlit)

## Features

- Upload one or more `.docx` templates containing `{{PLACEHOLDER}}` fields.
- Auto-detect placeholders in body text, tables, nested tables, headers/footers, and text boxes/floating shapes.
- Date fields use a calendar picker.
- Dates are inserted in `MM/DD/YYYY` format.
- Blank fields become blank in the generated document.
- Generate completed Word files while preserving surrounding template formatting.
- Download one Word file directly or multiple Word files as a ZIP.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Community Cloud

Commit only these files:

```text
app.py
requirements.txt
```

Do **not** add `packages.txt`. This version has no PDF conversion and does not require LibreOffice or any Linux/APT dependency.

## Date fields

Any placeholder whose name contains `DATE` uses the Streamlit calendar picker. The selected date is written as:

```text
MM/DD/YYYY
```

Example:

```text
09/09/2026
```

## Download

This version intentionally downloads **Word (.docx) only**. If multiple templates are generated, the app also provides a ZIP containing all completed Word files.
