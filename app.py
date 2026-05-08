import streamlit as st
import pandas as pd
import pdf2image
import tempfile
import os
import io
import re
from paddleocr import PaddleOCR

st.set_page_config(page_title="Relévé to Excel", layout="wide")
st.title("📄 Relévé Bancaire Scanné → Excel")
st.write("Mamaky PDF scanné, manala entête/pied de page, ary manavaka ny mouvement amin'ny daty.")

# Initialisation OCR (frantsay satria matetika amin'ny teny frantsay ny relévé)
@st.cache_resource
def load_ocr():
    return PaddleOCR(use_angle_cls=True, lang='fr', show_log=False)

ocr = load_ocr()

uploaded_file = st.file_uploader("📤 Ampidiro ny PDF scanné", type=["pdf"])

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded_file.read())
        tmp_path = tmp.name

    with st.spinner("⏳ Mamaky ny PDF... (mety haka 30s-2min arakaraka ny isan'ny pejy)"):
        try:
            images = pdf2image.convert_from_path(tmp_path, dpi=200)
            all_lines = []
            for img in images:
                result = ocr.ocr(img, cls=True)
                if result:
                    for line in result:
                        if line:
                            for word in line:
                                all_lines.append(word[1][0])
        except Exception as e:
            st.error(f"❌ Olana tamin'ny famakiana PDF: {e}")
            all_lines = []

    # Sivana: Tazonina fotsiny ny andalana manomboka amin'ny daty
    date_pattern = r'^\d{2}[/\.-]\d{2}[/\.-]\d{2,4}'
    mouvements = [line.strip() for line in all_lines if re.match(date_pattern, line)]

    if mouvements:
        data = []
        for line in mouvements:
            date_match = re.match(r'^(\d{2}[/\.-]\d{2}[/\.-]\d{2,4})\s*(.*)', line)
            if date_match:
                date = date_match.group(1)
                rest = date_match.group(2).strip()
                
                # Mitady vola any amin'ny farany (misaraka amin'ny habaka, misy , na .)
                amount_match = re.search(r'([0-9]{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?)\s*$', rest)
                if amount_match:
                    amount = amount_match.group(1)
                    libelle = rest[:rest.rfind(amount)].strip()
                    data.append({"Date": date, "Libellé": libelle, "Montant": amount})
                else:
                    data.append({"Date": date, "Libellé": rest, "Montant": ""})
        
        df = pd.DataFrame(data)
        st.success(f"✅ {len(df)} mouvement hita")
        st.dataframe(df, use_container_width=True)
        
        # Fanondranana Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        st.download_button(
            label="📥 Download Excel",
            data=output.getvalue(),
            file_name="mouvement_bancaire.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ Tsy nisy mouvement hita. Hamarino fa manomboka amin'ny daty (oh: 01/01/2024) ny andalana rehetra tianao alaina.")
    
    os.unlink(tmp_path)