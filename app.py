import streamlit as st
import pandas as pd
from pdf2image import convert_from_bytes
import pytesseract
import re
from io import BytesIO, StringIO
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="Relévé Bancaire to Excel", page_icon="📄", layout="wide")

st.title("📄 Convertisseur Relévé Bancaire → Excel")
st.write("Universel - mahay manavaka Débit/Crédit amin'ny position colonne")

# Initialisation session state
if 'df_result' not in st.session_state:
    st.session_state.df_result = None
if 'df_validated' not in st.session_state:
    st.session_state.df_validated = None
if 'debug_text' not in st.session_state:
    st.session_state.debug_text = ""

uploaded_file = st.file_uploader("Safidio ny fisie PDF", type=['pdf'])

# Bouton hanombohana ny lecture
if uploaded_file is not None and st.session_state.df_result is None:
    if st.button("🚀 Manomboka ny lecture", type="primary"):
        with st.spinner("Andalam-pamakiana ny PDF..."):
            images = convert_from_bytes(uploaded_file.read())
            
            all_rows = []
            debug_text = ""
            progress_bar = st.progress(0)
            
            date_pattern = re.compile(r'^\d{2}[/\.\-]\d{2}[/\.\-]\d{2,4}$')
            amount_pattern = re.compile(r'^\d{1,3}(?:[\s\.]\d{3})*,\d{2}$')
            
            for page_num, image in enumerate(images):
                # Mampiasa image_to_data mba hahafantarana ny POSITION X an'ny teny tsirairay
                data = pytesseract.image_to_data(
                    image,
                    lang='fra',
                    config='--psm 6 -c preserve_interword_spaces=1',
                    output_type=pytesseract.Output.DICT
                )
                
                img_width = image.width
                
                # Manangona ny teny isaky ny ligne (line_num)
                lines_dict = {}
                for i in range(len(data['text'])):
                    text = data['text'][i].strip()
                    if not text:
                        continue
                    
                    line_key = (data['block_num'][i], data['par_num'][i], data['line_num'][i])
                    if line_key not in lines_dict:
                        lines_dict[line_key] = []
                    
                    lines_dict[line_key].append({
                        'text': text,
                        'left': data['left'][i],
                        'width': data['width'][i],
                        'right': data['left'][i] + data['width'][i]
                    })
                
                # Mandalo ny ligne tsirairay
                for line_key in sorted(lines_dict.keys()):
                    words = lines_dict[line_key]
                    line_text = ' '.join([w['text'] for w in words])
                    debug_text += line_text + "\n"
                    
                    # Mitady daty sy montants
                    dates = []
                    amounts = []
                    libelle_words = []
                    
                    for w in words:
                        if date_pattern.match(w['text']):
                            dates.append(w)
                        elif amount_pattern.match(w['text']):
                            amounts.append(w)
                        else:
                            libelle_words.append(w['text'])
                    
                    # Raha tsy misy daty dia ligne fanampiny (suite an'ny libellé taloha)
                    if not dates:
                        # Mampiana ny libellé amin'ny ligne taloha
                        if all_rows and not amounts:
                            extra = ' '.join([w['text'] for w in words])
                            all_rows[-1]['Libelle ou Operation'] += ' ' + extra
                        continue
                    
                    date_op = dates[0]['text'] if len(dates) >= 1 else ''
                    date_val = dates[1]['text'] if len(dates) >= 2 else ''
                    
                    libelle = ' '.join(libelle_words).strip()
                    libelle = re.sub(r'\s+', ' ', libelle)
                    
                    debit = ''
                    credit = ''
                    
                    # Famantarana Débit/Crédit amin'ny POSITION X
                    # Crédit matetika eo amin'ny faran'ny pejy (>80% ny sakany)
                    # Débit eo afovoany (60-80%)
                    if amounts:
                        for amt in amounts:
                            center_x = amt['left'] + amt['width'] / 2
                            position_ratio = center_x / img_width
                            
                            if position_ratio > 0.82:
                                # Colonne Crédit
                                credit = amt['text']
                            else:
                                # Colonne Débit
                                debit = amt['text']
                    
                    all_rows.append({
                        'Date': date_op,
                        'Date de valeur': date_val,
                        'Libelle ou Operation': libelle,
                        'Debits': debit,
                        'Credits': credit
                    })
                
                progress_bar.progress((page_num + 1) / len(images))
            
            st.session_state.df_result = pd.DataFrame(all_rows)
            st.session_state.df_validated = st.session_state.df_result.copy()
            st.session_state.debug_text = debug_text
            st.rerun()

# Aseho ny vokatra raha efa misy
if st.session_state.df_result is not None:
    df = st.session_state.df_result
    
    st.success(f"✅ Vita ny lecture: {len(df)} ligne hita")
    
    st.markdown("---")
    st.subheader("📝 Modification (azonao ovaina mivantana)")
    st.info("Manaova modification dia tsindrio ny bouton **'Modification terminée'** ambany rehefa vita")
    
    # Editor (tsy mi-recharger satria session state)
    edited_df = st.data_editor(
        df,
        use_container_width=True,
        num_rows="dynamic",
        key="editor"
    )
    
    # Bouton "Modification terminée"
    col_btn1, col_btn2 = st.columns([1, 4])
    with col_btn1:
        if st.button("✅ Modification terminée", type="primary"):
            st.session_state.df_validated = edited_df.copy()
            st.success("Voatahiry ny modification!")
    with col_btn2:
        if st.button("🔄 Reset (lecture vaovao)"):
            st.session_state.df_result = None
            st.session_state.df_validated = None
            st.rerun()
    
    st.markdown("---")
    
    # Asehoy ny TOTAL amin'ny version validated
    if st.session_state.df_validated is not None:
        df_final = st.session_state.df_validated
        
        def parse_amount(val):
            if not val or pd.isna(val):
                return 0.0
            try:
                val_str = str(val).replace(' ', '').replace('.', '').replace(',', '.')
                return float(val_str)
            except:
                return 0.0
        
        df_calc = df_final.copy()
        df_calc['_debit_num'] = df_calc['Debits'].apply(parse_amount)
        df_calc['_credit_num'] = df_calc['Credits'].apply(parse_amount)
        
        total_debit = df_calc['_debit_num'].sum()
        total_credit = df_calc['_credit_num'].sum()
        solde = total_credit - total_debit
        
        st.subheader("💰 Total (après validation)")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Nb Mouvements", len(df_final))
        with col2:
            st.metric("Total Débits", f"{total_debit:,.2f}".replace(',', ' ').replace('.', ','))
        with col3:
            st.metric("Total Crédits", f"{total_credit:,.2f}".replace(',', ' ').replace('.', ','))
        with col4:
            st.metric("Solde", f"{solde:,.2f}".replace(',', ' ').replace('.', ','))
        
        if abs(solde) < 0.01:
            st.success("✅ Mifanaraka tsara ny compte (Débits = Crédits)")
        else:
            st.warning(f"⚠️ Tsy mifanaraka: misy écart {solde:,.2f}")
        
        st.markdown("---")
        st.subheader("📥 Export")
        
        col_a, col_b = st.columns(2)
        
        # Export Excel
        with col_a:
            output = BytesIO()
            
            excel_df = df_final.copy()
            excel_df['Debits'] = excel_df['Debits'].apply(parse_amount)
            excel_df['Credits'] = excel_df['Credits'].apply(parse_amount)
            
            def parse_date(d):
                if not d or pd.isna(d):
                    return None
                try:
                    return pd.to_datetime(d, dayfirst=True, errors='coerce')
                except:
                    return None
            
            excel_df['Date'] = excel_df['Date'].apply(parse_date)
            excel_df['Date de valeur'] = excel_df['Date de valeur'].apply(parse_date)
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                excel_df.to_excel(writer, index=False, sheet_name='Mouvements')
                
                worksheet = writer.sheets['Mouvements']
                
                header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
                header_font = Font(bold=True, color='FFFFFF')
                
                for cell in worksheet[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal='center')
                
                for row in worksheet.iter_rows(min_row=2, max_col=2):
                    for cell in row:
                        cell.number_format = 'DD/MM/YYYY'
                
                for row in worksheet.iter_rows(min_row=2, min_col=4, max_col=5):
                    for cell in row:
                        cell.number_format = '#,##0.00'
                
                worksheet.column_dimensions['A'].width = 12
                worksheet.column_dimensions['B'].width = 15
                worksheet.column_dimensions['C'].width = 50
                worksheet.column_dimensions['D'].width = 15
                worksheet.column_dimensions['E'].width = 15
                
                last_row = len(excel_df) + 2
                worksheet.cell(row=last_row, column=3, value='TOTAL').font = Font(bold=True)
                cell_d = worksheet.cell(row=last_row, column=4, value=total_debit)
                cell_d.number_format = '#,##0.00'
                cell_d.font = Font(bold=True)
                cell_c = worksheet.cell(row=last_row, column=5, value=total_credit)
                cell_c.number_format = '#,##0.00'
                cell_c.font = Font(bold=True)
            
            st.download_button(
                label="⬇️ Download Excel (.xlsx)",
                data=output.getvalue(),
                file_name='mouvement_bancaire.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True
            )
        
        # Export CSV
        with col_b:
            csv_buffer = StringIO()
            df_final.to_csv(csv_buffer, index=False, sep=';', encoding='utf-8-sig')
            
            st.download_button(
                label="⬇️ Download CSV (séparateur: ;)",
                data=csv_buffer.getvalue().encode('utf-8-sig'),
                file_name='mouvement_bancaire.csv',
                mime='text/csv',
                use_container_width=True
            )
    
    with st.expander("🔍 Lahatsoratra voavaky (Debug)"):
        st.text(st.session_state.debug_text[:5000])
