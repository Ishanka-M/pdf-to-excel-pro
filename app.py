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

    rows = []  # list of dicts; one dict per material/size row on the label

    # ── Order No. & internal seq number ──────────────────────────────────────
    order_match = re.search(r'Order No\.:(\S+)\s+(\d+)', text)
    order_no  = order_match.group(1) if order_match else ""
    seq_no    = order_match.group(2) if order_match else ""

    # ── Destination code (CMUS, etc.) ────────────────────────────────────────
    dest_match = re.search(r'Factory I/O:\s*\n(\S+)', text)
    destination = dest_match.group(1) if dest_match else ""

    # ── Ship From ─────────────────────────────────────────────────────────────
    # Format: "Ship From: <ID> Ship To:..."  then address lines follow
    ship_from_id_match = re.search(r'Ship From:\s*(\S+)', text)
    ship_from_id = ship_from_id_match.group(1) if ship_from_id_match else ""

    # Address lines before "Order No."
    address_block_match = re.search(
        r'Ship From:.*?\n(.*?)Order No\.',
        text, re.DOTALL
    )
    if address_block_match:
        addr_lines = address_block_match.group(1).strip().splitlines()
        # Left side only (right side has Ship-To address mixed in first lines)
        # First two lines are "Block 33…" repeated; last is "Loluwagoda, Mirigama"
        ship_from_addr = ", ".join(l.strip() for l in addr_lines if l.strip())
    else:
        ship_from_addr = ""

    ship_from = f"{ship_from_id} – {ship_from_addr}" if ship_from_id else ship_from_addr

    # ── Ship To ───────────────────────────────────────────────────────────────
    # The PDF is two-column; pdfplumber merges columns into one text stream.
    # The Ship To name is reliably on the "Ship To:" header line.
    # The address lines (10 Emery St / Bethlehem PA 18015) appear interleaved
    # with Ship From address lines, so we extract them separately.
    ship_to_header = re.search(r'Ship To:(.+)', text)
    to_name = clean_text(ship_to_header.group(1)) if ship_to_header else ""

    to_addr_match = re.search(
        r'Export Processing Zone\s+([\w\d].*?)\n.*?Bethlehem',
        text, re.DOTALL
    )
    to_street = clean_text(to_addr_match.group(1)) if to_addr_match else ""
    to_city_match = re.search(r'(Bethlehem PA \d+)', text)
    to_city = clean_text(to_city_match.group(1)) if to_city_match else ""

    ship_to = ", ".join(filter(None, [to_name, to_street, to_city]))

    # ── Carton Number & total cartons ─────────────────────────────────────────
    carton_match = re.search(
        r'CARTON NUMBER\s+(\d+)\s+of\s+(\d+)',
        text, re.IGNORECASE
    )
    carton_no    = carton_match.group(1) if carton_match else ""
    total_cartons = carton_match.group(2) if carton_match else ""
    carton_label = f"{carton_no} of {total_cartons}" if carton_no else ""

    # ── Carton Total (overall qty in this carton) ─────────────────────────────
    carton_total_match = re.search(r'CARTON TOTAL\s+(\d+)', text, re.IGNORECASE)
    carton_total = carton_total_match.group(1) if carton_total_match else ""

    # ── Label Total ───────────────────────────────────────────────────────────
    label_total_match = re.search(r'LABEL TOTAL\s+(\d+)', text, re.IGNORECASE)
    label_total = label_total_match.group(1) if label_total_match else ""

    # ── SSCC barcode ──────────────────────────────────────────────────────────
    # Printed as "(00) 0 0194698 005205193 2" – strip spaces & parens for raw digits
    sscc_raw_match = re.search(r'\(00\)([\d\s]+)', text)
    if sscc_raw_match:
        # Clean up: remove newlines/trailing page markers like "1/1"
        raw = sscc_raw_match.group(0).split('\n')[0].strip()
        sscc_display = raw                                    # e.g. "(00) 0 0194698 005205193 2"
        sscc_digits  = re.sub(r'[^\d]', '', raw)             # digits only, no parens/spaces
    else:
        sscc_digits  = ""
        sscc_display = ""

    # ── Material / Size / Quantity rows ──────────────────────────────────────
    # Everything between "Material # Size Quantity" header and "LABEL TOTAL"
    table_match = re.search(
        r'Material\s*#\s+Size\s+Quantity\s*\n(.*?)LABEL TOTAL',
        text, re.DOTALL | re.IGNORECASE
    )
    if table_match:
        table_text = table_match.group(1).strip()
        for line in table_text.splitlines():
            line = line.strip()
            if not line:
                continue
            # Each line: "<material_num>  <size>  <qty>"
            parts = line.split()
            if len(parts) >= 3:
                material = parts[0]
                size     = parts[1]
                qty      = parts[2]
            elif len(parts) == 2:
                material = parts[0]
                size     = parts[1]
                qty      = ""
            else:
                continue

            rows.append({
                "Order No."       : order_no,
                "Seq No."         : seq_no,
                "Destination"     : destination,
                "Ship From"       : ship_from,
                "Ship To"         : ship_to,
                "Material #"      : material,
                "Size"            : size,
                "Quantity"        : qty,
                "Label Total"     : label_total,
                "Carton Total"    : carton_total,
                "Carton No."      : carton_label,
                "SSCC (display)"  : sscc_display,
                "SSCC (digits)"   : sscc_digits,
            })

    # Fallback: if table parsing failed, still return one row with metadata
    if not rows:
        rows.append({
            "Order No."       : order_no,
            "Seq No."         : seq_no,
            "Destination"     : destination,
            "Ship From"       : ship_from,
            "Ship To"         : ship_to,
            "Material #"      : "",
            "Size"            : "",
            "Quantity"        : "",
            "Label Total"     : label_total,
            "Carton Total"    : carton_total,
            "Carton No."      : carton_label,
            "SSCC (display)"  : sscc_display,
            "SSCC (digits)"   : sscc_digits,
        })

    return rows


# ── Column display order ───────────────────────────────────────────────────────
COLUMN_ORDER = [
    "Order No.", "Seq No.", "Destination",
    "Ship From", "Ship To",
    "Material #", "Size", "Quantity",
    "Label Total", "Carton Total", "Carton No.",
    "SSCC (display)", "SSCC (digits)",
]


if uploaded_file is not None:
    try:
        with st.spinner("දත්ත කියවමින් පවතී..."):
            all_rows = []

            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        page_rows = extract_label_data(text)
                        all_rows.extend(page_rows)

            if all_rows:
                df = pd.DataFrame(all_rows)
                # Ensure all expected columns exist
                for col in COLUMN_ORDER:
                    if col not in df.columns:
                        df[col] = ""
                df = df[COLUMN_ORDER]

                total_labels = df["Carton No."].nunique()
                st.success(
                    f"✅ ලේබල් {total_labels} ක දත්ත (පේළි {len(df)}) සාර්ථකව හඳුනා ගන්නා ලදී!"
                )
                st.dataframe(df, use_container_width=True)

                # ── Build Excel output ────────────────────────────────────────
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Shipping_Data')

                    workbook  = writer.book
                    worksheet = writer.sheets['Shipping_Data']

                    header_fmt = workbook.add_format({
                        'bold'      : True,
                        'bg_color'  : '#1E1E1E',
                        'font_color': 'white',
                        'border'    : 1,
                        'align'     : 'center',
                        'valign'    : 'vcenter',
                    })
                    data_fmt = workbook.add_format({
                        'border' : 1,
                        'valign' : 'vcenter',
                    })

                    col_widths = {
                        "Order No."      : 16,
                        "Seq No."        : 8,
                        "Destination"    : 13,
                        "Ship From"      : 40,
                        "Ship To"        : 35,
                        "Material #"     : 18,
                        "Size"           : 8,
                        "Quantity"       : 10,
                        "Label Total"    : 12,
                        "Carton Total"   : 12,
                        "Carton No."     : 13,
                        "SSCC (display)" : 30,
                        "SSCC (digits)"  : 25,
                    }

                    for col_num, col_name in enumerate(df.columns):
                        worksheet.write(0, col_num, col_name, header_fmt)
                        worksheet.set_column(
                            col_num, col_num,
                            col_widths.get(col_name, 20),
                            data_fmt
                        )

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
