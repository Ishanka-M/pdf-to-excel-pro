import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# Interface එකේ පෙනුම සැකසීම
st.set_page_config(page_title="AI Shipping Label Extractor", layout="wide")

# ශීර්ෂය සහ නිර්මාණකරුගේ නම
st.title("📊 AI Shipping Label to Excel Converter")
st.subheader("Developed by Ishanka Madusanka")
st.markdown("---")

uploaded_file = st.file_uploader("ඔබේ PDF ලේබල් ගොනුව මෙතැනට Upload කරන්න", type="pdf")

def clean_text(text):
    if text:
        return re.sub(r'\s+', ' ', text).strip()
    return ""

def extract_advanced_logic(text):
    """සියලුම දත්ත ක්ෂේත්‍ර වෙන් කර හඳුනා ගැනීමේ තර්කය"""
    
    # SSCC අංකය (පහළ ඇති දිගු අංකය)
    sscc_match = re.search(r'(\d{18,20})', text)
    sscc = sscc_match.group(1) if sscc_match else "N/A"

    # SHIP FROM සහ SHIP TO (ලිපිනයන් පේළි කිහිපයක ඇති බැවින් ඒවා වෙන් කර ගැනීම)
    ship_from = re.search(r'SHIP FROM:\s*(.*?)(?=SHIP TO:)', text, re.DOTALL)
    ship_to = re.search(r'SHIP TO:\s*(.*?)(?=PO#:)', text, re.DOTALL)
    
    # අනෙකුත් දත්ත
    po = re.search(r'PO#:\s*(\S+)', text)
    style = re.search(r'STYLE#:\s*(\S+)', text)
    asin = re.search(r'ASIN#:\s*(\S+)', text)
    
    # Item Description
    item_desc = re.search(r'ITEM DESC:\s*(.*?)(?=ASIN#|UPC:|QTY:|$)', text, re.DOTALL)
    
    qty = re.search(r'QTY:\s*(\d+)', text)
    carton = re.search(r'CARTON#:\s*(\d+\s*of\s*\d+)', text)

    return {
        "SSCC Number": sscc,
        "SHIP FROM": clean_text(ship_from.group(1)) if ship_from else "",
        "SHIP TO": clean_text(ship_to.group(1)) if ship_to else "",
        "PO #": po.group(1) if po else "",
        "STYLE #": style.group(1) if style else "",
        "ASIN #": asin.group(1) if asin else "",
        "ITEM DESCRIPTION": clean_text(item_desc.group(1)) if item_desc else "",
        "QTY": qty.group(1) if qty else "",
        "CARTON #": carton.group(1) if carton else ""
    }

if uploaded_file is not None:
    try:
        with st.spinner("දත්ත කියවමින් පවතී..."):
            extracted_list = []
            
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        data_row = extract_advanced_logic(text)
                        extracted_list.append(data_row)

            if extracted_list:
                df = pd.DataFrame(extracted_list)
                
                # තීරු පිළිවෙල සැකසීම
                order = ["SSCC Number", "SHIP FROM", "SHIP TO", "PO #", "STYLE #", "ASIN #", "ITEM DESCRIPTION", "QTY", "CARTON #"]
                df = df[order]

                st.success(f"ලේබල් {len(extracted_list)} ක දත්ත සාර්ථකව හඳුනා ගන්නා ලදී!")
                st.dataframe(df, use_container_width=True)

                # Excel ගොනුව සකස් කිරීම
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Shipping_Data')
                    
                    workbook = writer.book
                    worksheet = writer.sheets['Shipping_Data']
                    
                    # Headers වල පෙනුම
                    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1E1E1E', 'font_color': 'white', 'border': 1})
                    for col_num, value in enumerate(df.columns.values):
                        worksheet.write(0, col_num, value, header_fmt)
                        worksheet.set_column(col_num, col_num, 25) # තීරු වල පළල

                st.download_button(
                    label="📥 Download Master Excel File",
                    data=output.getvalue(),
                    file_name="Shipping_Master_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
        st.error(f"දෝෂයක් සිදුවිය: {e}")

# Footer එකේ නිර්මාණකරුගේ නම පෙන්වීම
st.markdown("---")
st.caption("© 2024 AI Shipping Tool | Developed by **Ishanka Madusanka**")
