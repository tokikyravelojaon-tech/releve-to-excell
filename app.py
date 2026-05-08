import streamlit as st
import pandas as pd
from pdf2image import convert_from_bytes
import pytesseract
import re
from io import BytesIO

st.set_page_config(page_title="Relévé Bancaire to Excel", page_icon="📄", layout="wide")

st.title("📄 Convertisseur Relévé Bancaire → Excel")
st.write("Format: Date | Date de valeur | Libellé | Débits | Crédits")

uploaded_file = st.file_uploader("Safidio ny fisie PDF", type=['pdf'])

if uploaded_file is not None:
    with st.spinner("Andalam-pamakiana ny PDF..."):
        images = convert_from_bytes(uploaded_file.read())
        all_text = ""
        progress_bar = st.progress(0)
        
        for i, image in enumerate(images):
            # Mampiasa config manokana mba hitazonana ny espaces (colonne)
            text = pytesseract.image_to_string(
                image, 
                lang='fra',
                config='--psm 6 -c preserve_interword_spaces=1'
            )
            all_text += text + "\n"
            progress_bar.progress((i + 1) / len(images))
        
        st.success(f"✅ Vita ny famakiana {len(images)} pejy!")
        
        all_lines = all_text.split('\n')
        
        # Pattern: ligne manomboka amin'ny daty (format JJ/MM na JJ/MM/AAAA)
        date_start_pattern = re.compile(r'^\s*(\d{2}[/\.\-]\d{2}(?:[/\.\-]\d{2,4})?)')
        amount_pattern = re.compile(r'(\d{1,3}(?:[\s\.]\d{3})*,\d{2})')
        
        mouvements = []
        
        for line in all_lines:
            # Esorina ny espaces betsaka eo am-piandohana fa tazonina ny eo afovoany
            line_clean = line.rstrip()
            if not line_clean.strip():
                continue
            
            # Mijery raha manomboka amin'ny daty
            date_match = date_start_pattern.match(line_clean)
            if not date_match:
                continue
            
            date_op = date_match.group(1)
            rest = line_clean[date_match.end():].strip()
            
            # Mitady ny daty faharoa (Date de valeur) - optionnel
            date_val_match = re.match(r'^(\d{2}[/\.\-]\d{2}(?:[/\.\-]\d{2,4})?)\s+', rest)
            date_val = ''
            if date_val_match:
                date_val = date_val_match.group(1)
                rest = rest[date_val_match.end():].strip()
            
            # Mitady ny montants rehetra ao amin'ny ligne
            amounts = amount_pattern.findall(rest)
            
            debit = ''
            credit = ''
            libelle = rest
            
            if amounts:
                # Famantarana Débit/Crédit amin'ny POSITION (colonne) ao amin'ny ligne
                # Eo amin'ny relévé bancaire: Débit eo afovoany, Crédit any an-tsisiny
                
                last_amount = amounts[-1]
                last_pos = rest.rfind(last_amount)
                
                # Esorina ny montant amin'ny libellé
                libelle = rest[:last_pos].strip()
                # Esorina ihany koa ny montants hafa amin'ny libellé
                for amt in amounts:
                    libelle = libelle.replace(amt, '').strip()
                
                # Famantarana arakaraka ny position amin'ny ligne
                # Raha lavitra be ny montant (>70% an'ny ligne) → Crédit
                # Raha eo afovoany → Débit
                line_length = len(rest)
                position_ratio = last_pos / line_length if line_length > 0 else 0
                
                # Famantarana fanampiny amin'ny mots-clés
                libelle_lower = libelle.lower()
                credit_keywords = ['vir recu', 'virement recu', 'versement', 'remise', 
                                   'depot', 'salaire', 'remboursement', 'vir sepa recu']
                debit_keywords = ['retrait', 'frais', 'commission', 'cotisation', 'prelv', 
                                 'prelevement', 'paiement', 'cheque', 'achat cb', 
                                 'vir sepa emis', 'facture']
                
                is_credit = any(kw in libelle_lower for kw in credit_keywords)
                is_debit = any(kw in libelle_lower for kw in debit_keywords)
                
                if is_credit:
                    credit = last_amount
                elif is_debit:
                    debit = last_amount
                else:
                    # Default: arakaraka ny position
                    if position_ratio > 0.75:
                        credit = last_amount
                    else:
                        debit = last_amount
            
            # Ampidirina raha tsy banga ny libellé
            if libelle and len(libelle) > 2:
                mouvements.append({
                    'Date': date_op,
                    'Date de valeur': date_val if date_val else date_op,
                    'Libellé ou Opération': libelle,
                    'Débits': debit,
                    'Crédits': credit
                })
        
        if mouvements:
            df = pd.DataFrame(mouvements)
            st.subheader(f"📊 Mouvement hita: {len(df)}")
            st.dataframe(df, use_container_width=True)
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total", len(df))
            with col2:
                st.metric("Débits", len(df[df['Débits'] != '']))
            with col3:
                st.metric("Crédits", len(df[df['Crédits'] != '']))
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Mouvements')
            
            st.download_button(
                label="⬇️ Download Excel",
                data=output.getvalue(),
                file_name='mouvement_bancaire.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            st.warning("⚠️ Tsy nisy mouvement hita.")
        
        # Debug toujours affiché
        with st.expander("🔍 Lahatsoratra voavaky (Debug) - Alefaso ahy ny screenshot"):
            st.text(all_text[:3000])
