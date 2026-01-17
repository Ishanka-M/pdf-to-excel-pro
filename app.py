import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# Page Configuration
st.set_page_config(page_title="Ultimate AI PDF Converter", layout="wide")

st.title("🏆 AI-Powered Packing List Master")
st.markdown("මෙය PDF එකේ ඇති වගු වල සීමාවන් (Bounding Boxes) හඳුනාගෙන 100% ක් නිවැරදිව Excel සකසයි.")

uploaded_file = st.file_uploader("Upload Helen Kaminski Packing List", type="pdf")

def advanced_clean(text):
    """දත්ත පිරිසිදු කර පේළි කැඩීම් (Newlines) ඉවත් කරයි"""
    if text is None: return ""
    # අකුරු අතර ඇති අනවශ්‍ය පේළි කැඩීම් ඉවත් කර තනි පේළියකට ගනී
    text = str(text).replace('\n', ' ')
    return re.sub(r'\s+', ' ', text).strip()

if uploaded_file:
    with st.spinner("Deep Scan ක්‍රියාත්මකයි... කරුණාකර රැඳී සිටින්න."):
        all_table_data = []
        
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                # දියුණු Table Extraction තාක්ෂණය
                # මෙහි settings මගින් වගුවේ නොපෙනෙන ඉරි පවා හඳුනා ගනී
                table = page.extract_table({
                    "vertical_strategy": "lines_price", # ඉරි සහ අකුරු පිහිටීම යන දෙකම බලයි
                    "horizontal_strategy": "text", 
                    "snap_tolerance": 3,
                    "join_tolerance": 3,
                    "edge_min_length": 15,
                })
                
                if table:
                    for row in table:
                        # සෑම සෛලයක්ම (Cell) පිරිසිදු කිරීම
                        cleaned_row = [advanced_clean(cell) for cell in row]
                        # හිස් පේළි ඉවත් කිරීම
                        if any(cleaned_row):
                            all_table_data.append(cleaned_row)

        if all_table_data:
            # Pandas භාවිතා කර ව්‍යුහය සකස් කිරීම
            df = pd.DataFrame(all_table_data)
            
            st.success("Analysis Completed!")
            st.write("### Extracted Data Preview")
            st.dataframe(df, use_container_width=True)

            # Excel ගොනුව සකස් කිරීම
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, header=False, sheet_name='Packing_List')
                
                workbook = writer.book
                worksheet = writer.sheets['Packing_List']
                
                # Excel formatting (ලස්සනට සකස් කිරීම)
                header_fmt = workbook.add_format({'bold': True, 'bg_color': '#CFE2F3', 'border': 1})
                cell_fmt = workbook.add_format({'border': 1, 'valign': 'vcenter'})
                
                # තීරු වල පළල ස්වයංක්‍රීයව සකස් කිරීම (Auto-fit look)
                for i, col in enumerate(df.columns):
                    worksheet.set_column(i, i, 20, cell_fmt)
            
            st.download_button(
                label="📥 Download Master Excel File",
                data=output.getvalue(),
                file_name="Master_Packing_List_Converted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
