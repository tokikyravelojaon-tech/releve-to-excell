import streamlit as st
import pandas as pd
from pdf2image import convert_from_bytes
import pytesseract
import re
from io import BytesIO, StringIO
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="Relévé Bancaire to Excel", page_icon="📄", layout="wide")

st.title("📄 Convertisseur Relévé Bancaire → Excel")

# Init session state
if 'df_result' not in st.session_state:
    st.session_state.df_result = None
if 'df_validated' not in st.session_state:
    st.session_state.df_validated = None
if 'debug_text' not in st.session_state:
    st.session_state.debug_text = ""

uploaded_file = st.file_uploader("Safidio ny fisie PDF", type=['pdf'])

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
                data = pytesseract.image_to_data(
                    image,
                    lang='fra',
                    config='--psm 6 -c preserve_interword_spaces=1',
                    output_type=pytesseract.Output.DICT
                )
                
                img_width = image.width
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
                        'width': data['width'][i]
                    })
                
                for line_key in sorted(lines_dict.keys()):
                    words = lines_dict[line_key]
                    line_text = ' '.join([w['text'] for w in words])
                    debug_text += line_text + "\n"
                    
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
                    
                    if not dates:
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
                    
                    if amounts:
                        for amt in amounts:
                            center_x = amt['left'] + amt['width'] / 2
                            position_ratio = center_x / img_width
                            if position_ratio > 0.82:
                                credit = amt['text']
                            else:
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

if st.session_state.df_result is not None:
    df = st.session_state.df_result
    
    st.success(f"✅ Vita ny lecture: {len(df)} ligne hita")
    
    st.markdown("---")
    st.subheader("📝 Modification (azonao ovaina mivantana)")
    st.info("Manaova modification dia tsindrio **'Modification terminée'** rehefa vita")
    
    edited_df = st.data_editor(
        df,
        use_container_width=True,
        num_rows="dynamic",
        key="editor"
    )
    
    col_btn1, col_btn2 = st.columns([1, 4])
    with col_btn1:
        if st.button("✅ Modification terminée", type="primary", use_container_width=True):
            st.session_state.df_validated = edited_df.copy().reset_index(drop=True)
            st.session_state.df_result = edited_df.copy().reset_index(drop=True)
            st.success("Voatahiry sy nokajiana indray ny Total!")
            st.rerun()
    with col_btn2:
        if st.button("🔄 Reset (PDF vaovao)"):
            st.session_state.df_result = None
            st.session_state.df_validated = None
            st.session_state.debug_text = ""
            st.rerun()
    
    st.markdown("---")
    
    if st.session_state.df_validated is not None:
        df_final = st.session_state.df_validated.copy()
        
        def parse_amount(val):
            if val is None or pd.isna(val) or str(val).strip() == '':
                return 0.0
            try:
                val_str = str(val).strip().replace(' ', '').replace('.', '').replace(',', '.')
                return float(val_str)
            except:
                return 0.0
        
        # RECALCUL TANTERAKA isaky ny rerun
        df_calc = df_final.copy()
        df_calc['_d'] = df_calc['Debits'].apply(parse_amount)
        df_calc['_c'] = df_calc['Credits'].apply(parse_amount)
        
        total_debit = float(df_calc['_d'].sum())
        total_credit = float(df_calc['_c'].sum())
        solde = total_credit - total_debit
        
        def fmt(n):
            return f"{n:,.2f}".replace(',', ' ').replace('.', ',')
        
        st.subheader("💰 Total (recalculé)")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown(f"<div style='background:#f0f2f6;padding:15px;border-radius:8px;text-align:center;'><div style='color:#666;font-size:14px;'>Nb Mouvements</div><div style='font-size:26px;font-weight:bold;color:#1F4E78;'>{len(df_final)}</div></div>", unsafe_allow_html=True)
        with col2:
            st.markdown(f"<div style='background:#f8d7da;padding:15px;border-radius:8px;text-align:center;'><div style='color:#666;font-size:14px;'>Total Débits</div><div style='font-size:26px;font-weight:bold;color:#dc3545;'>{fmt(total_debit)}</div></div>", unsafe_allow_html=True)
        with col3:
            st.markdown(f"<div style='background:#d4edda;padding:15px;border-radius:8px;text-align:center;'><div style='color:#666;font-size:14px;'>Total Crédits</div><div style='font-size:26px;font-weight:bold;color:#28a745;'>{fmt(total_credit)}</div></div>", unsafe_allow_html=True)
        with col4:
            color = "#28a745" if abs(solde) < 0.01 else "#fd7e14"
            st.markdown(f"<div style='background:#e2e3e5;padding:15px;border-radius:8px;text-align:center;'><div style='color:#666;font-size:14px;'>Solde (C - D)</div><div style='font-size:26px;font-weight:bold;color:{color};'>{fmt(solde)}</div></div>", unsafe_allow_html=True)
        
        st.markdown("")
        
        if abs(solde) < 0.01:
            st.success(f"✅ Mifanaraka tsara ny compte: Débits = Crédits = {fmt(total_debit)}")
        else:
            st.warning(f"⚠️ Écart: {fmt(solde)} | Débits: {fmt(total_debit)} | Crédits: {fmt(total_credit)}")
        
        st.markdown("---")
        st.subheader("📥 Export")
        
        col_a, col_b = st.columns(2)
        
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
                ws = writer.sheets['Mouvements']
                
                hf = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
                hfont = Font(bold=True, color='FFFFFF')
                for cell in ws[1]:
                    cell.fill = hf
                    cell.font = hfont
                    cell.alignment = Alignment(horizontal='center')
                
                for row in ws.iter_rows(min_row=2, max_col=2):
                    for cell in row:
                        cell.number_format = 'DD/MM/YYYY'
                
                for row in ws.iter_rows(min_row=2, min_col=4, max_col=5):
                    for cell in row:
                        cell.number_format = '#,##0.00'
                
                ws.column_dimensions['A'].width = 12
                ws.column_dimensions['B'].width = 15
                ws.column_dimensions['C'].width = 50
                ws.column_dimensions['D'].width = 15
                ws.column_dimensions['E'].width = 15
                
                lr = len(excel_df) + 2
                ws.cell(row=lr, column=3, value='TOTAL').font = Font(bold=True)
                cd = ws.cell(row=lr, column=4, value=total_debit)
                cd.number_format = '#,##0.00'
                cd.font = Font(bold=True)
                cc = ws.cell(row=lr, column=5, value=total_credit)
                cc.number_format = '#,##0.00'
                cc.font = Font(bold=True)
            
            st.download_button(
                label="⬇️ Download Excel (.xlsx)",
                data=output.getvalue(),
                file_name='mouvement_bancaire.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True
            )
        
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
