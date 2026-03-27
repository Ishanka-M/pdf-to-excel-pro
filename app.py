import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="AI Shipping Label Extractor", layout="wide")

st.title("📊 AI Shipping Label to Excel Converter")
st.subheader("Developed by Ishanka Madusanka")
st.markdown("---")

uploaded_file = st.file_uploader("ඔබේ PDF ලේබල් ගොනුව මෙතැනට Upload කරන්න", type="pdf")


def clean_text(text):
    if text:
        return re.sub(r'\s+', ' ', text).strip()
    return ""


def extract_label_data(text):
    """
    Extract all fields from the CMUS / Club Monaco shipping label format.

    Raw text layout (pdfplumber output):
      Ship From: 0202400082 Ship To:CARTPA-Club Monaco US
      Block 33, Export Processing Zone 10 Emery St
      Block 33, Export Processing Zone, Bethlehem PA 18015
      Loluwagoda, Mirigama
      Order No.:CM01-005728 261
      Customer Order No.: Customer Department
      Factory I/O:
      CMUS
      Material # Size Quantity
      295100901100 XS 50          <- one or more rows
      LABEL TOTAL 50
      CARTON NUMBER 1 of 13 CARTON TOTAL 50
      (00) 0 0194698 005205193 2
      1/1
    """

    rows = []

    # Order No. & internal seq number
    order_match = re.search(r'Order No\.:(\S+)\s+(\d+)', text)
    order_no  = order_match.group(1) if order_match else ""
    seq_no    = order_match.group(2) if order_match else ""

    # Destination code (CMUS, etc.)
    dest_match = re.search(r'Factory I/O:\s*\n(\S+)', text)
    destination = dest_match.group(1) if dest_match else ""

    # Ship From
    ship_from_id_match = re.search(r'Ship From:\s*(\S+)', text)
    ship_from_id = ship_from_id_match.group(1) if ship_from_id_match else ""
    address_block_match = re.search(r'Ship From:.*?\n(.*?)Order No\.', text, re.DOTALL)
    if address_block_match:
        addr_lines = address_block_match.group(1).strip().splitlines()
        ship_from_addr = ", ".join(l.strip() for l in addr_lines if l.strip())
    else:
        ship_from_addr = ""
    ship_from = f"{ship_from_id} – {ship_from_addr}" if ship_from_id else ship_from_addr

    # Ship To (two-column PDF; name on header line, street/city extracted separately)
    ship_to_header = re.search(r'Ship To:(.+)', text)
    to_name = clean_text(ship_to_header.group(1)) if ship_to_header else ""
    to_addr_match = re.search(r'Export Processing Zone\s+([\w\d].*?)\n.*?Bethlehem', text, re.DOTALL)
    to_street = clean_text(to_addr_match.group(1)) if to_addr_match else ""
    to_city_match = re.search(r'(Bethlehem PA \d+)', text)
    to_city = clean_text(to_city_match.group(1)) if to_city_match else ""
    ship_to = ", ".join(filter(None, [to_name, to_street, to_city]))

    # Carton Number
    carton_match = re.search(r'CARTON NUMBER\s+(\d+)\s+of\s+(\d+)', text, re.IGNORECASE)
    carton_no     = carton_match.group(1) if carton_match else ""
    total_cartons = carton_match.group(2) if carton_match else ""
    carton_label  = f"{carton_no} of {total_cartons}" if carton_no else ""

    carton_total_match = re.search(r'CARTON TOTAL\s+(\d+)', text, re.IGNORECASE)
    carton_total = carton_total_match.group(1) if carton_total_match else ""

    label_total_match = re.search(r'LABEL TOTAL\s+(\d+)', text, re.IGNORECASE)
    label_total = label_total_match.group(1) if label_total_match else ""

    # SSCC barcode
    sscc_raw_match = re.search(r'\(00\)([\d\s]+)', text)
    if sscc_raw_match:
        raw          = sscc_raw_match.group(0).split('\n')[0].strip()
        sscc_display = raw
        sscc_digits  = re.sub(r'[^\d]', '', raw)
    else:
        sscc_display = sscc_digits = ""

    # Material / Size / Quantity rows
    table_match = re.search(
        r'Material\s*#\s+Size\s+Quantity\s*\n(.*?)LABEL TOTAL',
        text, re.DOTALL | re.IGNORECASE
    )
    if table_match:
        for line in table_match.group(1).strip().splitlines():
            line = line.strip()
            if not line:
                continue
            parts = line.split()
            if len(parts) >= 3:
                material, size, qty = parts[0], parts[1], parts[2]
            elif len(parts) == 2:
                material, size, qty = parts[0], parts[1], ""
            else:
                continue
            rows.append({
                "Order No."      : order_no,
                "Seq No."        : seq_no,
                "Destination"    : destination,
                "Ship From"      : ship_from,
                "Ship To"        : ship_to,
                "Material #"     : material,
                "Size"           : size,
                "Quantity"       : qty,
                "Label Total"    : label_total,
                "Carton Total"   : carton_total,
                "Carton No."     : carton_label,
                "SSCC (display)" : sscc_display,
                "SSCC (digits)"  : sscc_digits,
            })

    if not rows:
        rows.append({
            "Order No."      : order_no,
            "Seq No."        : seq_no,
            "Destination"    : destination,
            "Ship From"      : ship_from,
            "Ship To"        : ship_to,
            "Material #"     : "",
            "Size"           : "",
            "Quantity"       : "",
            "Label Total"    : label_total,
            "Carton Total"   : carton_total,
            "Carton No."     : carton_label,
            "SSCC (display)" : sscc_display,
            "SSCC (digits)"  : sscc_digits,
        })

    return rows


def merge_sscc_groups(df):
    """
    Merge rows that share the same SSCC (= same physical carton with multiple sizes).
    - Size      → "L/S/XS"
    - Size Detail → "L:18/S:18/XS:6"
    - Quantity  → sum of all sizes
    Single-size cartons pass through unchanged (Size Detail left empty).
    """
    def _merge(group):
        first = group.iloc[0].copy()
        if len(group) > 1:
            first["Size"]        = "/".join(group["Size"].tolist())
            first["Size Detail"] = "/".join(
                f'{r["Size"]}:{r["Quantity"]}' for _, r in group.iterrows()
            )
            first["Quantity"]    = str(sum(
                int(q) for q in group["Quantity"] if str(q).isdigit()
            ))
        else:
            first["Size Detail"] = ""
        return first

    return (
        df.groupby("SSCC (digits)", sort=False, group_keys=False)
          .apply(_merge)
          .reset_index(drop=True)
    )


COLUMN_ORDER = [
    "Order No.", "Seq No.", "Destination",
    "Ship From", "Ship To",
    "Material #", "Size", "Size Detail", "Quantity",
    "Label Total", "Carton Total", "Carton No.",
    "SSCC (display)", "SSCC (digits)",
]

COL_WIDTHS = {
    "Order No."      : 16,
    "Seq No."        : 8,
    "Destination"    : 13,
    "Ship From"      : 40,
    "Ship To"        : 35,
    "Material #"     : 18,
    "Size"           : 10,
    "Size Detail"    : 22,
    "Quantity"       : 10,
    "Label Total"    : 12,
    "Carton Total"   : 12,
    "Carton No."     : 13,
    "SSCC (display)" : 30,
    "SSCC (digits)"  : 25,
}


if uploaded_file is not None:
    try:
        with st.spinner("දත්ත කියවමින් පවතී..."):
            all_rows = []
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        all_rows.extend(extract_label_data(text))

            if all_rows:
                df = pd.DataFrame(all_rows)

                # Merge duplicate-SSCC rows (mixed-size cartons)
                df = merge_sscc_groups(df)

                # Ensure all columns present
                for col in COLUMN_ORDER:
                    if col not in df.columns:
                        df[col] = ""
                df = df[COLUMN_ORDER]

                total_labels = df["Carton No."].nunique()
                st.success(
                    f"✅ ලේබල් {total_labels} ක දත්ත (පේළි {len(df)}) සාර්ථකව හඳුනා ගන්නා ලදී!"
                )
                st.dataframe(df, use_container_width=True)

                # Build Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Shipping_Data')
                    workbook  = writer.book
                    worksheet = writer.sheets['Shipping_Data']

                    header_fmt = workbook.add_format({
                        'bold': True, 'bg_color': '#1E1E1E',
                        'font_color': 'white', 'border': 1,
                        'align': 'center', 'valign': 'vcenter',
                    })
                    data_fmt = workbook.add_format({'border': 1, 'valign': 'vcenter'})
                    mixed_fmt = workbook.add_format({
                        'border': 1, 'valign': 'vcenter',
                        'bg_color': '#FFF9C4',  # light yellow for mixed-size rows
                    })

                    for col_num, col_name in enumerate(df.columns):
                        worksheet.write(0, col_num, col_name, header_fmt)
                        worksheet.set_column(col_num, col_num, COL_WIDTHS.get(col_name, 20))

                    # Write data rows; highlight mixed-size cartons
                    for row_idx, row in df.iterrows():
                        fmt = mixed_fmt if "/" in str(row.get("Size", "")) else data_fmt
                        for col_idx, col_name in enumerate(df.columns):
                            worksheet.write(row_idx + 1, col_idx, row[col_name], fmt)

                    worksheet.set_row(0, 20)

                st.download_button(
                    label="📥 Download Master Excel File",
                    data=output.getvalue(),
                    file_name="Shipping_Master_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("PDF එකෙන් දත්ත හඳුනා ගත නොහැකි විය.")

    except Exception as e:
        st.error(f"දෝෂයක් සිදුවිය: {e}")
        raise

st.markdown("---")
st.caption("© 2024 AI Shipping Tool | Developed by **Ishanka Madusanka**")
