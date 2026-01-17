import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="AI PDF to Excel Master", layout="wide")

st.title("📊 AI-Driven Packing List Converter")
st.markdown("මෙම පද්ධතිය Helen Kaminski Packing List වැනි සංකීර්ණ ලේඛන සඳහාම විශේෂිතව නිපදවා ඇත.")

uploaded_file = st.file_uploader("ඔබේ PDF ගොනුව මෙතැනට Upload කරන්න", type="pdf")

def smart_clean(text):
    if text is None: return ""
    # පේළි කැඩීම් සහ අනවශ්‍ය හිස්තැන් ඉවත් කර තනි පේළියකට ගනී
    return re.sub(r'\s+', ' ', str(text)).strip()

if uploaded_file is not None:
    try:
        with st.spinner("දත්ත විශ්ලේෂණය කරමින් පවතී..."):
            all_data = []
            
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    # 'Lattice' තාක්ෂණය: වගුවේ නොපෙනෙන කෝෂ (Cells) හඳුනා ගනී
                    table = page.extract_table({
                        "vertical_strategy": "text",
                        "horizontal_strategy": "text",
                        "snap_tolerance": 5,
                        "join_tolerance": 5,
                    })
                    
                    if table:
                        for row in table:
                            # සෑම සෛලයක්ම පිරිසිදු කිරීම
                            cleaned_row = [smart_clean(cell) for cell in row]
                            if any(cleaned_row): # හිස් පේළි ඉවත් කිරීම
                                all_data.append(cleaned_row)

            if all_data:
                df = pd.DataFrame(all_data)
                
                # Excel formatting සහ Download බටන් එක
                st.success("සාර්ථකයි!")
                st.dataframe(df, use_container_width=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, header=False, sheet_name='Data')
                    
                    # තීරු වල පළල ස්වයංක්‍රීයව සැකසීම
                    worksheet = writer.sheets['Data']
                    for i, _ in enumerate(df.columns):
                        worksheet.set_column(i, i, 22)

                st.download_button(
                    label="📥 Download Master Excel File",
                    data=output.getvalue(),
                    file_name="Converted_Packing_List.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
        st.error(f"Error එකක් සිදු විය: {str(e)}")
        st.info("කරුණාකර requirements.txt ගොනුව නිවැරදි දැයි පරීක්ෂා කරන්න.")
