import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="SSCC Data Extractor", layout="wide")

st.title("📦 SSCC-Based Label Converter")
st.markdown("මෙම පද්ධතිය මගින් SSCC අංකය පදනම් කරගෙන ඊට අදාළ දත්ත වෙන් කර ලබා දෙයි.")

uploaded_file = st.file_uploader("ලේබල් සහිත PDF එක Upload කරන්න", type="pdf")

def extract_sscc_logic(text):
    """SSCC අංකය මුල් කරගෙන අනෙක් දත්ත ගලපන ශ්‍රිතය"""
    
    # SSCC අංකය සෙවීම (ලේබලයේ පහළ ඇති ඉලක්කම් 18-20 ක අංකය)
    sscc_match = re.search(r'(\d{18,20})', text)
    sscc = sscc_match.group(1) if sscc_match else "හමු නොවීය"

    # අනෙකුත් දත්ත සොයා ගැනීම
    po = re.search(r'PO#:\s*(\S+)', text)
    style = re.search(r'STYLE#:\s*(\S+)', text)
    
    # Item Description එක සාමාන්‍යයෙන් පේළි කිහිපයක් විය හැක
    item_desc = re.search(r'ITEM DESC:\s*(.*?)(?=ASIN#|UPC:|QTY:|$)', text, re.DOTALL)
    
    qty = re.search(r'QTY:\s*(\d+)', text)
    carton = re.search(r'CARTON#:\s*(\d+\s*of\s*\d+)', text)

    return {
        "SSCC (Serial Shipping Container Code)": sscc,
        "PO #": po.group(1) if po else "",
        "STYLE #": style.group(1) if style else "",
        "ITEM DESCRIPTION": item_desc.group(1).replace('\n', ' ').strip() if item_desc else "",
        "QTY": qty.group(1) if qty else "",
        "CARTON #": carton.group(1) if carton else ""
    }

if uploaded_file is not None:
    try:
        with st.spinner("දත්ත විශ්ලේෂණය කරමින්..."):
            extracted_list = []
            
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        # සෑම ලේබලයකම දත්ත SSCC එකට අනුව ගැලපීම
                        data_row = extract_sscc_logic(text)
                        extracted_list.append(data_row)

            if extracted_list:
                # දත්ත වගුවක් ලෙස සැකසීම
                df = pd.DataFrame(extracted_list)
                
                # Column පිළිවෙල සැකසීම (SSCC මුලට)
                cols = ["SSCC (Serial Shipping Container Code)", "PO #", "STYLE #", "ITEM DESCRIPTION", "QTY", "CARTON #"]
                df = df[cols]

                st.success("දත්ත සාර්ථකව වෙන් කරගන්නා ලදී!")
                st.dataframe(df, use_container_width=True)

                # Excel Download Option
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='SSCC_Report')
                    
                    # Excel පෙනුම සැකසීම
                    workbook = writer.book
                    worksheet = writer.sheets['SSCC_Report']
                    header_format = workbook.add_format({'bold': True, 'bg_color': '#212F3D', 'font_color': 'white', 'border': 1})
                    
                    for col_num, value in enumerate(df.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                        worksheet.set_column(col_num, col_num, 25)

                st.download_button(
                    label="📥 Download SSCC Master Excel",
                    data=output.getvalue(),
                    file_name="SSCC_Shipping_Data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
        st.error(f"Error: {e}")
