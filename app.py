import streamlit as st
import pdfplumber
import pandas as pd
import io

# පිටුවේ පෙනුම සකස් කිරීම
st.set_page_config(page_title="Packing List Converter", layout="wide")

st.title("📑 Professional PDF to Excel Converter")
st.info("Upload your Helen Kaminski Packing List to convert.")

uploaded_file = st.file_uploader("Choose your PDF file", type="pdf")

if uploaded_file is not None:
    try:
        all_rows = []
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                # දත්ත වගු ලෙස වෙන් කරගැනීම
                table = page.extract_table({
                    "vertical_strategy": "text", 
                    "horizontal_strategy": "text",
                    "snap_tolerance": 4,
                })
                if table:
                    all_rows.extend(table)
        
        if all_rows:
            df = pd.DataFrame(all_rows)
            st.write("### Preview of Converted Data")
            st.dataframe(df.head(10)) # පළමු පේළි 10 පෙන්වයි

            # Excel එක සෑදීම
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, header=False)
            
            st.download_button(
                label="📥 Download Full Excel File",
                data=output.getvalue(),
                file_name="Converted_List.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No tabular data detected. Please check the PDF quality.")
            
    except Exception as e:
        st.error(f"An error occurred: {e}")
