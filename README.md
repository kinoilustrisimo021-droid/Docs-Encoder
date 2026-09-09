# Docx Template Filler (Streamlit)

## Paano patakbuhin

1. I-install ang requirements:
   ```
   pip install -r requirements.txt
   ```
2. Patakbuhin ang app:
   ```
   streamlit run app.py
   ```
3. Buksan sa browser (auto-open, or http://localhost:8501)

## Paano gamitin

1. **Upload** — i-drag/drop ang isa o higit pang `.docx` na may `{{PLACEHOLDER}}` fields
   (e.g. `{{Client_Name}}`, `{{BRAND}}`, `{{UNIT_DESCRIPTION}}`).
2. Awtomatikong ma-detect lahat ng unique `{{...}}` fields sa file(s) — kasama na
   ang mga nasa loob ng tables, headers/footers, **at text boxes / floating shapes**
   (madalas nasa text box ang mga logo/endorsement sections sa mga bank forms,
   kaya sinisigurado ng app na masi-scan din yun).
3. **Fill out the form** — isang beses lang i-type ang value; kung parehong field
   ang lumalabas sa dalawang template (hal. Client_Name), isang input lang ang
   kailangan, applied na sa lahat.
4. **I-generate at i-download agad** — pagka-submit ng form, direktang ma-generate
   ang final na `.docx` (o `.zip` kung marami) at lalabas na yung download button —
   wala nang hiwalay na "Review" page.

## Deployment (Streamlit Community Cloud)

1. I-push ang `app.py` at `requirements.txt` sa isang GitHub repo.
2. Sa https://share.streamlit.io , i-connect ang repo, piliin ang `app.py` bilang
   entry point, deploy.
3. Multiple users can use it at the same time — walang shared/global state,
   session-scoped lahat (`st.session_state`) at in-memory lang ang processing
   (walang temp files sa disk), kaya safe at magaan kahit maraming user sabay-sabay.

## Mahalagang paalala

- Original formatting (fonts, bold, tables, spacing, logos sa text boxes) ng
  template ay **hindi nagbabago** — {{placeholders}} lang ang pinapalitan,
  kahit saan pa man ito nasa loob ng document.
- Kung may bagong template na ipapa-upload ng users mo na may ibang field
  names (basta naka-`{{LIKE_THIS}}` format), automatic na siyang made-detect —
  walang kailangang i-edit sa code.
- Pagkatapos i-download, may button pang "Edit fields & regenerate" kung
  gusto mong ayusin lang ang typo nang hindi na-uupload ulit ang file.
