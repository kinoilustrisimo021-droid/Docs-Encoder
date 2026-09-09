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
   ang mga nasa loob ng tables at headers/footers.
3. **Fill out the form** — isang beses lang i-type ang value; kung parehong field
   ang lumalabas sa dalawang template (hal. Client_Name), isang input lang ang
   kailangan, applied na sa lahat.
4. **Review** — makikita ang lahat ng ininput bago i-generate; may text preview
   din per document.
5. **Download** — kukunin ang final `.docx` (kung isa lang ang na-upload) o
   isang `.zip` (kung marami).

## Deployment (Streamlit Community Cloud)

1. I-push ang `app.py` at `requirements.txt` sa isang GitHub repo.
2. Sa https://share.streamlit.io , i-connect ang repo, piliin ang `app.py` bilang
   entry point, deploy.
3. Multiple users can use it at the same time — walang shared/global state,
   session-scoped lahat (`st.session_state`) at in-memory lang ang processing
   (walang temp files sa disk), kaya safe at magaan kahit maraming user sabay-sabay.

## Mahalagang paalala

- Original formatting (fonts, bold, tables, spacing) ng template ay
  **hindi nagbabago** — {{placeholders}} lang ang pinapalitan.
- Kung may bagong template na ipapa-upload ng users mo na may ibang field
  names (basta naka-`{{LIKE_THIS}}` format), automatic na siyang made-detect —
  walang kailangang i-edit sa code.
- Ang preview sa "Review" step ay plain text lang (para mabilis/magaan) —
  ang actual downloaded `.docx` ang may buong Word formatting.
