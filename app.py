import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="Shipping Label to Excel", layout="wide")

st.title("📦 Shipping Label Data Extractor")
st.markdown("Brandix/Amazon Label වැනි ලේඛන වලින් දත්ත නිවැරදිව Excel වලට ලබා ගැනීමට මෙය භාවිතා කරන්න.")

uploaded_file = st.file_uploader("ලේබල් සහිත PDF ගොනුව Upload කරන්න", type="pdf")

def extract_label_data(text):
    """PDF පෙළෙහි ඇති දත්ත Regex මගින් වෙන් කර හඳුනා ගනී"""
    data = {}
    
    # එක් එක් දත්ත ක්ෂේත්‍රය හඳුනා ගැනීමට patterns භාවිතා කිරීම
    data['PO#'] = re.search(r'PO#:\s*(.*)', text)
    data['STYLE#'] = re.search(r'STYLE#:\s*(.*)', text)
    data['ITEM DESC'] = re.search(r'ITEM DESC:\s*(.*?)(?=ASIN#|UPC:|$)', text, re.DOTALL)
    data['ASIN#'] = re.search(r'ASIN#:\s*(.*)', text)
    data['UPC'] = re.search(r'UPC:\s*(.*)', text)
    data['QTY'] = re.search(r'QTY:\s*(\d+)', text)
    data['CARTON#'] = re.search(r'CARTON#:\s*(.*)', text)
    data['Country of Origin'] = re.search(r'Country Of Origin\s*(.*)', text)
    
    # SSCC Barcode අංකය (පහළ ඇති දිගු අංකය)
    sscc_match = re.search(r'(\d{18,20})$', text.strip())
    data['SSCC'] = sscc_match.group(1) if sscc_match else ""

    # දත්ත පිරිසිදු කර නිවැරදි අගය පමණක් ලබා ගැනීම
    return {k: (v.group(1).strip() if hasattr(v, 'group') and v else "") for k, v in data.items()}

if uploaded_file is not None:
    try:
        with st.spinner("ලේබල් කියවමින් පවතී..."):
            all_labels = []
            
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    # මුළු පිටුවේම ඇති Text එක ලබා ගැනීම
                    text = page.extract_text()
                    if text:
                        label_info = extract_label_data(text)
                        all_labels.append(label_info)

            if all_labels:
                df = pd.DataFrame(all_labels)
                
                # පෙනුම සැකසීම
                st.success(f"ලේබල් {len(all_labels)} ක් සාර්ථකව හඳුනා ගන්නා ලදී!")
                st.dataframe(df, use_container_width=True)

                # Excel එක සෑදීම
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Labels')
                    
                    # Excel Formatting
                    workbook = writer.book
                    worksheet = writer.sheets['Labels']
                    header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
                    
                    for col_num, value in enumerate(df.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                        worksheet.set_column(col_num, col_num, 20)

                st.download_button(
                    label="📥 Download Excel File",
                    data=output.getvalue(),
                    file_name="Label_Data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("කිසිදු දත්තයක් හඳුනා ගැනීමට නොහැකි විය. කරුණාකර PDF එකේ ගුණාත්මකභාවය පරීක්ෂා කරන්න.")

    except Exception as e:
        st.error(f"දෝෂයක් සිදුවිය: {e}")
