import streamlit as st
import pandas as pd
from pdf2image import convert_from_bytes
import pytesseract
import re
from io import BytesIO, StringIO
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="Relévé Bancaire to Excel", page_icon="📄", layout="wide")

st.title("📄 Convertisseur Relévé Bancaire → Excel")
st.write("Maka izay rehetra hita ao amin'ny PDF (universel)")

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
        
        st.success(f"Vita ny famakiana {len(images)} pejy!")
        
        all_lines = all_text.split('\n')
        
        date_pattern = re.compile(r'(\d{2}[/\.\-]\d{2}(?:[/\.\-]\d{2,4})?)')
        amount_pattern = re.compile(r'(\d{1,3}(?:[\s\.]\d{3})*,\d{2}|\d+\.\d{2})')
        
        mouvements = []
        
        for line in all_lines:
            line_clean = line.rstrip()
            if not line_clean.strip():
                continue
            
            if not re.match(r'^\s*\d{2}[/\.\-]\d{2}', line_clean):
                continue
            
            dates_found = date_pattern.findall(line_clean)
            
            if not dates_found:
                continue
            
            date_op = dates_found[0] if len(dates_found) >= 1 else ''
            date_val = dates_found[1] if len(dates_found) >= 2 else ''
            
            rest = line_clean
            for d in dates_found[:2]:
                rest = rest.replace(d, '', 1)
            
            amounts = amount_pattern.findall(rest)
            
            debit = ''
            credit = ''
            libelle = rest
            
            if amounts:
                for amt in amounts:
                    libelle = libelle.replace(amt, '', 1)
                
                if len(amounts) == 1:
                    pos = rest.rfind(amounts[0])
                    if pos / len(rest) > 0.7:
                        credit = amounts[0]
                    else:
                        debit = amounts[0]
                elif len(amounts) >= 2:
                    debit = amounts[0]
            
            libelle = re.sub(r'\s+', ' ', libelle).strip()
            
            mouvements.append({
                'Date': date_op,
                'Date de valeur': date_val,
                'Libelle ou Operation': libelle,
                'Debits': debit,
                'Credits': credit
            })
        
        if mouvements:
            df = pd.DataFrame(mouvements)
            
            st.subheader(f"Mouvement hita: {len(df)}")
            st.info("Azonao ovaina mivantana eto ambany ny vokatra alohan'ny export")
            
            edited_df = st.data_editor(
                df, 
                use_container_width=True, 
                num_rows="dynamic"
            )
            
            def parse_amount(val):
                if not val or pd.isna(val):
                    return 0.0
                try:
                    val_str = str(val).replace(' ', '').replace('.', '').replace(',', '.')
                    return float(val_str)
                except:
                    return 0.0
            
            edited_df['_debit_num'] = edited_df['Debits'].apply(parse_amount)
            edited_df['_credit_num'] = edited_df['Credits'].apply(parse_amount)
            
            total_debit = edited_df['_debit_num'].sum()
            total_credit = edited_df['_credit_num'].sum()
            solde = total_credit - total_debit
            
            st.subheader("Total")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Nb Mouvements", len(edited_df))
            with col2:
                st.metric("Total Debits", f"{total_debit:,.2f}")
            with col3:
                st.metric("Total Credits", f"{total_credit:,.2f}")
            with col4:
                st.metric("Solde", f"{solde:,.2f}")
            
            if abs(solde) < 0.01:
                st.success("Mifanaraka tsara ny compte (Debits = Credits)")
            else:
                st.warning(f"Tsy mifanaraka: misy ecart {solde:,.2f}")
            
            export_df = edited_df.drop(columns=['_debit_num', '_credit_num'])
            
            st.markdown("---")
            st.subheader("Export")
            
            col_a, col_b = st.columns(2)
            
            with col_a:
                output = BytesIO()
                
                excel_df = export_df.copy()
                excel_df['Debits'] = excel_df['Debits'].apply(parse_amount)
                excel_df['Credits'] = excel_df['Credits'].apply(parse_amount)
                
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
                    
                    workbook = writer.book
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
                    label="Download Excel (.xlsx)",
                    data=output.getvalue(),
                    file_name='mouvement_bancaire.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True
                )
            
            with col_b:
                csv_buffer = StringIO()
                export_df.to_csv(csv_buffer, index=False, sep=';', encoding='utf-8-sig')
                
                st.download_button(
                    label="Download CSV (separateur: ;)",
                    data=csv_buffer.getvalue().encode('utf-8-sig'),
                    file_name='mouvement_bancaire.csv',
                    mime='text/csv',
                    use_container_width=True
                )
        else:
            st.warning("Tsy nisy mouvement hita.")
        
        with st.expander("Lahatsoratra voavaky (Debug)"):
            st.text(all_text[:3000])
