import streamlit as st
import pandas as pd
from pdf2image import convert_from_bytes
import pytesseract
import re
from io import BytesIO

st.set_page_config(page_title="Relévé Bancaire to Excel", page_icon="📄", layout="wide")

st.title("📄 Convertisseur Relévé Bancaire → Excel")
st.write("Ampidiro ny PDF scanné dia ny mouvement ihany no halaina.")

uploaded_file = st.file_uploader("Safidio ny fisie PDF", type=['pdf'])

if uploaded_file is not None:
    with st.spinner("Andalam-pamakiana ny PDF..."):
        images = convert_from_bytes(uploaded_file.read())
        all_lines = []
        progress_bar = st.progress(0)
        
        for i, image in enumerate(images):
            text = pytesseract.image_to_string(image, lang='fra')
            lines = text.split('\n')
            all_lines.extend(lines)
            progress_bar.progress((i + 1) / len(images))
        
        st.success(f"Vita ny famakiana {len(images)} pejy!")
        
        date_pattern = r'^(\d{2}[/\.\-]\d{2}[/\.\-]\d{2,4})'
        mouvements = []
        
        for line in all_lines:
            line = line.strip()
            if re.match(date_pattern, line):
                match = re.match(r'^(\d{2}[/\.\-]\d{2}[/\.\-]\d{2,4})\s+(.*)', line)
                if match:
                    date_val = match.group(1)
                    rest = match.group(2)
                    amount_match = re.search(r'(\d{1,3}(?:[\s\.]\d{3})*,\d{2}|\d+\.\d{2})\s*$', rest)
                    if amount_match:
                        amount = amount_match.group(1)
                        libelle = rest[:amount_match.start()].strip()
                        mouvements.append({'Date': date_val, 'Libelle': libelle, 'Montant': amount})
                    else:
                        mouvements.append({'Date': date_val, 'Libelle': rest, 'Montant': ''})
        
        if mouvements:
            df = pd.DataFrame(mouvements)
            st.subheader(f"Mouvement hita: {len(df)}")
            st.dataframe(df, use_container_width=True)
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Mouvements')
            
            st.download_button(
                label="Download Excel",
                data=output.getvalue(),
                file_name='mouvement_bancaire.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            st.warning("Tsy nisy mouvement hita.")
