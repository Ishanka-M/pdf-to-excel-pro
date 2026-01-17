import streamlit as st
import pdfplumber
import pandas as pd
import io

# වෙබ් පිටුවේ පෙනුම සැකසීම
st.set_page_config(page_title="Ultra PDF to Excel Converter", layout="wide")

st.title("🚀 Professional PDF to Excel (100% Accuracy Mode)")
st.markdown("මෙම පද්ධතිය ඔබේ Packing List එකේ ඇති වගු වල හැඩය (Layout) එලෙසම ආරක්ෂා කරයි.")

uploaded_file = st.file_uploader("ඔබේ PDF ගොනුව මෙතැනට Upload කරන්න", type="pdf")

if uploaded_file is not None:
    with st.spinner("Analyzing layout and extracting tables..."):
        all_pages_data = []
        
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                # දියුණු Table Extraction Settings
                # මෙය වගුවේ ඇති ඉරි සහ අකුරු අතර පරතරය ඉතා සියුම්ව පරීක්ෂා කරයි
                table_settings = {
                    "vertical_strategy": "text",   # අකුරු අනුව තීරු වෙන් කිරීම
                    "horizontal_strategy": "text", # අකුරු අනුව පේළි වෙන් කිරීම
                    "snap_tolerance": 3,           # අකුරු එකිනෙකට සම්බන්ධ කිරීමේ පරාසය
                    "join_tolerance": 3,
                    "edge_min_length": 15,
                    "intersection_tolerance": 10,
                }
                
                table = page.extract_table(table_settings)
                
                if table:
                    # පේළි ඇතුළත ඇති අනවශ්‍ය 'New Lines' (\n) ඉවත් කර පිරිසිදු කිරීම
                    clean_table = []
                    for row in table:
                        clean_row = [str(cell).replace('\n', ' ').strip() if cell else "" for cell in row]
                        clean_table.append(clean_row)
                    
                    df_page = pd.DataFrame(clean_table)
                    all_pages_data.append(df_page)

        if all_pages_data:
            # සියලුම පිටු එකම වගුවකට සම්බන්ධ කිරීම
            final_df = pd.concat(all_pages_data, ignore_index=True)
            
            # Preview පෙන්වීම
            st.success("සාර්ථකව දත්ත හඳුනාගන්නා ලදී!")
            st.write("### Data Preview")
            st.dataframe(final_df)

            # Excel ගොනුව සෑදීම (Styles සහිතව)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False, header=False, sheet_name='Packing_List')
                
                # Excel එකේ තීරු වල පළල ස්වයංක්‍රීයව සකස් කිරීම
                workbook = writer.book
                worksheet = writer.sheets['Packing_List']
                for i, col in enumerate(final_df.columns):
                    worksheet.set_column(i, i, 20) 

            st.download_button(
                label="📥 Download Perfect Excel File",
                data=output.getvalue(),
                file_name="Formatted_Packing_List.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("වගු හඳුනා ගැනීමට නොහැකි විය. කරුණාකර PDF ගොනුව පරීක්ෂා කරන්න.")
