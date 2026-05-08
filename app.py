import streamlit as st
import pandas as pd
from pdf2image import convert_from_bytes
import pytesseract
import re
from io import BytesIO, StringIO
from openpyxl.styles import NamedStyle, Font, PatternFill, Alignment

st.set_page_config(page_title="Relévé Bancaire to Excel", page_icon="📄", layout="wide")

st.title("📄 Convertisseur Relévé Bancaire → Excel")
st.write("Maka **izay rehetra hita** ao amin'ny PDF (universel - mety amin'ny banky rehetra)")

uploaded_file = st.file_uploader("Safidio ny fisie PDF", type=['pdf'])

if uploaded_file is not None:
    with st.spinner("Andalam-pamakiana ny PDF..."):
        images = convert_from_bytes(uploaded_file.read())
        all_text = ""
        progress_bar = st.progress(0)
        
        for i, image in enumerate(images):
            text = pytesseract.image_to_string(
                image, 
                lang='fra',
                config='--psm 6 -c preserve_interword_spaces=1'
            )
            all_text += text + "\n"
            progress_bar.progress((i + 1) / len(images))
        
        st.success(f"✅ Vita ny famakiana {len(images)} pejy!")
        
        all_lines = all_text.split('\n')
        
        # Pattern
        date_pattern = re.compile(r'(\d{2}[/\.\-]\d{2}(?:[/\.\-]\d{2,4})?)')
        amount_pattern = re.compile(r'(\d{1,3}(?:[\s\.]\d{3})*,\d{2}|\d+\.\d{2})')
        
        mouvements = []
        
        for line in all_lines:
            line_clean = line.rstrip()
            if not line_clean.strip():
                continue
            
            # Mila manomboka amin'ny daty ny ligne
            if not re.match(r'^\s*\d{2}[/\.\-]\d{2}', line_clean):
                continue
            
            # Mitady daty rehetra (maximum 2: Date sy Date de valeur)
            dates_found = date_pattern.findall(line_clean)
            
            if not dates_found:
                continue
            
            date_op = dates_found[0] if len(dates_found) >= 1 else ''
            date_val = dates_found[1] if len(dates_found) >= 2 else ''
            
            # Esorina ny daty hita amin'ny ligne mba ho voafaritra ny libellé
            rest = line_clean
            for d in dates_found[:2]:  # Esorina ny 2 daty voalohany ihany
                rest = rest.replace(d, '', 1)
            
            # Mitady ny montants
            amounts = amount_pattern.findall(rest)
            
            debit = ''
            credit = ''
            libelle = rest
            
            if amounts:
                # Esorina ny montants amin'ny libellé
                for amt in amounts:
                    libelle = libelle.replace(amt, '', 1)
                
                # Logique: 
                # - Raha 1 montant ihany → mametraka ao Débit (default)
                # - Raha 2 montants → voalohany = Débit, faharoa = Crédit (na solde)
                #   → Eto: voalohany = mouvement, faharoa = solde, ka mametraka eo @ position
                
                if len(amounts) == 1:
                    # Famantarana amin'ny position (raha lavitra dia Crédit)
                    pos = rest.rfind(amounts[0])
                    if pos / len(rest) > 0.7:
                        credit = amounts[0]
                    else:
                        debit = amounts[0]
                elif len(amounts) >= 2:
                    # 2 montants: voalohany = mouvement, faharoa = solde
                    # Mametraka ny mouvement ao @ Débit (azonao ovaina manually)
                    debit = amounts[0]
            
            # Manadio libellé (esorina espaces betsaka)
            libelle = re.sub(r'\s+', ' ', libelle).strip()
            
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
            st.info("💡 Azonao ovaina mivantana eto ambany ny vokatra alohan'ny export")
            
            # Tabilao azo ovaina (editable)
            edited_df = st.data_editor(
                df, 
                use_container_width=True, 
                num_rows="dynamic",
                column_config={
                    "Date": st.column_config.TextColumn("Date", width="small"),
                    "Date de valeur": st.column_config.TextColumn("Date de valeur", width="small"),
                    "Libellé ou Opération": st.column_config.TextColumn("Libellé", width="large"),
                    "Débits": st.column_config.TextColumn("Débits", width="small"),
                    "Crédits": st.column_config.TextColumn("Crédits", width="small"),
                }
            )
            
            # Calcul Total
            def parse_amount(val):
                if not val or pd.isna(val):
                    return 0.0
                try:
                    # Format français: 1 250,50 → 1250.50
                    val_str = str(val).replace(' ', '').replace('.', '').replace(',', '.')
                    return float(val_str)
                except:
                    return 0.0
            
            edited_df['_debit_num'] = edited_df['Débits'].apply(parse_amount)
            edited_df['_credit_num'] = edited_df['Crédits'].apply(parse_amount)
            
            total_debit = edited_df['_debit_num'].sum()
            total_credit = edited_df['_credit_num'].sum()
            solde = total_credit - total_debit
            
            # Affichage Total
            st.subheader("💰 Total")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Nb Mouvements", len(edited_df))
            with col2:
                st.metric("Total Débits", f"{total_debit:,.2f} €".replace(',', ' ').replace('.', ','))
            with col3:
                st.metric("Total Crédits", f"{total_credit:,.2f} €".replace(',', ' ').replace('.', ','))
            with col4:
                st.metric("Solde (C-D)", f"{solde:,.2f} €".replace(',', ' ').replace('.', ','),
                         delta="Équilibré" if abs(solde) < 0.01 else f"Écart: {solde:.2f}")
            
            # Vérification équilibre
            if abs(solde) < 0.01:
                st.success("✅ Mifanaraka tsara ny compte (Débits = Crédits)")
            else:
                st.warning(f"⚠️ Tsy mifanaraka: misy ecart {solde:,.2f} €")
            
            # Esorina ny colonne tampon alohan'ny export
            export_df = edited_df.drop(columns=['_debit_num', '_credit_num'])
            
            st.markdown("---")
            st.subheader("📥 Export")
            
            col_a, col_b = st.columns(2)
            
            # ============ EXPORT EXCEL avec formatage ============
            with col_a:
                output = BytesIO()
                
                # Avadika ho format tsara ny daty sy vola
                excel_df = export_df.copy()
                excel_df['Débits'] = excel_df['Débits'].apply(parse_amount)
                excel_df['Crédits'] = excel_df['Crédits'].apply(parse_amount)
                
                # Avadika daty
                def parse_date(d):
                    if not d:
                        return None
                    try:
                        return pd.to_datetime(d, dayfirst=True, errors='coerce')
                    except:
                        return None
                
                excel_df['Date'] = excel_df['Date'].apply(parse_date)
                excel_df['Date de valeur'] = excel_df['Date de valeur'].apply(parse_date)
                
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    excel_df.to_excel(writer, index=False, sheet_name='Mouvements')
                    
                    # Formatage
                    workbook = writer.book
                    worksheet = writer.sheets['Mouvements']
                    
                    # Header style
                    header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
                    header_font = Font(bold=True, color='FFFFFF')
                    
                    for cell in worksheet[1]:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = Alignment(horizontal='center')
                    
                    # Format date pour colonnes A et B
                    for row in worksheet.iter_rows(min_row=2, max_col=2):
                        for cell in row:
                            cell.number_format = 'DD/MM/YYYY'
                    
                    # Format monétaire pour colonnes D et E (Débits, Crédits)
                    for row in worksheet.iter_rows(min_row=2, min_col=4, max_col=5):
                        for cell in row:
                            cell.number_format = '#,##0.00 €'
                    
                    # Largeur colonne
                    worksheet.column_dimensions['A'].width = 12
                    worksheet.column_dimensions['B'].width = 15
                    worksheet.column_dimensions['C'].width = 50
                    worksheet.column_dimensions['D'].width = 15
                    worksheet.column_dimensions['E'].width = 15
                    
                    # Ligne Total eo amin'ny faran'ny tabilao
                    last_row = len(excel_df) + 2
                    worksheet.cell(row=last_row, column=3, value='TOTAL').font = Font(bold=True)
                    worksheet.cell(row=last_row, column=4, value=total_debit).number_format = '#,##0.00 €'
                    worksheet.cell(row=last_row, column=4).font = Font(bold=True)
                    worksheet.cell(row=last_row, column=5, value=total_credit).number_format = '#,##0.00 €'
                    worksheet.cell(row=last_row, column=5).font = Font(bold=True)
                
                st.download_button(
                    label="⬇️ Download Excel (.xlsx)",
                    data=output.getvalue(),
                    file_name='mouvement_bancaire.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True
                )
            
            # ============ EXPORT CSV avec ; ============
            with col_b:
                csv_buffer = StringIO()
                export_df.to_csv(csv_buffer, index=False, sep=';', encoding='utf-8-sig')
                
                st.download_button(
                    label="⬇️ Download CSV (séparateur: ;)",
                    data=csv_buffer.getvalue().encode('utf-8-sig'),
                    file_name='mouvement_bancaire.csv',
                    mime='text/csv',
                    use_container_width=True
                )
        else:
            st.warning("⚠️ Tsy nisy mouvement hita.")
        
        with st.expander("🔍 Lahatsoratra voavaky (Debug)"):
            st.text(all_text[:3000])import streamlit as st
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
