import streamlit as st
import pandas as pd
from pdf2image import convert_from_bytes
import pytesseract
import re
from io import BytesIO

st.set_page_config(page_title="Relévé Bancaire to Excel", page_icon="📄", layout="wide")

st.title("📄 Convertisseur Relévé Bancaire → Excel")
st.write("Ampidiro ny PDF scanné dia hivoaka amin'ny Excel ny mouvement (Date, Date de valeur, Libellé, Débits, Crédits).")

uploaded_file = st.file_uploader("Safidio ny fisie PDF", type=['pdf'])

# Fanazavana ny fomba famantarana Débit/Crédit
st.sidebar.header("⚙️ Paramètres")
detection_mode = st.sidebar.radio(
    "Fomba famantarana Débit/Crédit:",
    ["Auto (2 montants = Débit + Solde)", "Manuel par mots-clés", "Position (colonne)"]
)

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
        
        st.success(f"✅ Vita ny famakiana {len(images)} pejy!")
        
        # Pattern: andalana manomboka amin'ny daty ROA
        # Ohatra: "01/10/23 03/10/23 VIREMENT XXX 150 000,00 1 250 000,00"
        line_pattern = re.compile(
            r'^(\d{2}[/\.\-]\d{2}[/\.\-]\d{2,4})\s+'      # Date
            r'(\d{2}[/\.\-]\d{2}[/\.\-]\d{2,4})\s+'        # Date de valeur
            r'(.+)$'                                          # Reste (libellé + montants)
        )
        
        # Pattern montant: 1 250 000,00 na 1.250.000,00 na 1250000.00
        amount_pattern = re.compile(r'(\d{1,3}(?:[\s\.]\d{3})*,\d{2}|\d+\.\d{2})')
        
        # Mots-clés débit
        debit_keywords = ['retrait', 'frais', 'commission', 'agios', 'prelev', 'paiement', 
                         'cheque', 'virement emis', 'achat', 'tpe', 'gab']
        credit_keywords = ['versement', 'virement recu', 'remise', 'depot', 'salaire', 
                          'credit', 'remboursement']
        
        mouvements = []
        
        for line in all_lines:
            line = line.strip()
            match = line_pattern.match(line)
            
            if match:
                date_op = match.group(1)
                date_val = match.group(2)
                rest = match.group(3)
                
                # Mitady ny montants rehetra
                amounts = amount_pattern.findall(rest)
                
                debit = ''
                credit = ''
                libelle = rest
                
                if amounts:
                    # Esorina amin'ny libellé ny montants hita
                    for amt in amounts:
                        libelle = libelle.replace(amt, '')
                    libelle = libelle.strip()
                    
                    # Logique arakaraka ny mode
                    if detection_mode == "Auto (2 montants = Débit + Solde)":
                        # Raha 2 montants: voalohany = mouvement, faharoa = solde
                        # Mila mamaritra raha débit na crédit amin'ny mots-clés
                        montant = amounts[0]
                        libelle_lower = libelle.lower()
                        
                        if any(kw in libelle_lower for kw in credit_keywords):
                            credit = montant
                        elif any(kw in libelle_lower for kw in debit_keywords):
                            debit = montant
                        else:
                            # Default: débit (matetika)
                            debit = montant
                    
                    elif detection_mode == "Manuel par mots-clés":
                        montant = amounts[0]
                        libelle_lower = libelle.lower()
                        
                        if any(kw in libelle_lower for kw in credit_keywords):
                            credit = montant
                        else:
                            debit = montant
                    
                    elif detection_mode == "Position (colonne)":
                        # Raha 2 montants: voalohany débit, faharoa crédit
                        if len(amounts) >= 2:
                            debit = amounts[0]
                            credit = amounts[1]
                        else:
                            debit = amounts[0]
                
                mouvements.append({
                    'Date': date_op,
                    'Date de valeur': date_val,
                    'Libellé ou Opération': libelle,
                    'Débits': debit,
                    'Crédits': credit
                })
        
        if mouvements:
            df = pd.DataFrame(mouvements)
            st.subheader(f"📊 Mouvement hita: {len(df)}")
            st.dataframe(df, use_container_width=True)
            
            # Statistique kely
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Mouvement", len(df))
            with col2:
                debits_count = len(df[df['Débits'] != ''])
                st.metric("Débits", debits_count)
            with col3:
                credits_count = len(df[df['Crédits'] != ''])
                st.metric("Crédits", credits_count)
            
            # Export Excel
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
            st.warning("⚠️ Tsy nisy mouvement hita. Hamarino ny format-n'ny PDF.")
            
            with st.expander("🔍 Jereo ny lahatsoratra voavaky (Debug)"):
                st.text('\n'.join(all_lines[:50]))
