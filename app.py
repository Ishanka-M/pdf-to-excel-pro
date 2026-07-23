import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="AI Shipping Label Extractor", layout="wide")

st.title("📊 AI Shipping Label to Excel Converter")
st.subheader("Developed by Lakshan")
st.markdown("---")

uploaded_file = st.file_uploader("ඔබේ PDF ලේබල් ගොනුව මෙතැනට Upload කරන්න", type="pdf")


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────
def clean_text(text):
    if text:
        return re.sub(r'\s+', ' ', text).strip()
    return ""


def reverse_lines(text):
    if not text:
        return []
    return [line[::-1] for line in text.split('\n')]


def gtin_check_digit(d13):
    total = 0
    for i, ch in enumerate(reversed(d13)):
        total += int(ch) * (3 if i % 2 == 0 else 1)
    return (10 - (total % 10)) % 10


def detect_format(text):
    if not text:
        return None
    # අලුත් format එක හඳුනාගැනීම (Unichela / McDonough ආදිය)
    if "UNICHELA" in text or "NON-CONFORMING" in text or "MCDONOUGH" in text:
        return "unichela"
    if "Material #" in text or "Ship From:" in text:
        return "cmus"
    rev = " ".join(reverse_lines(text))
    if "CARTON" in rev and ("(01)" in rev or "STYLE/COLOR" in rev):
        return "carton"
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Parser 1 — Rotated GS1 carton label
# ─────────────────────────────────────────────────────────────────────────────
def parse_carton_label(text, page_no):
    lines = reverse_lines(text)
    joined = " ".join(lines)

    m = re.search(r'(\w+)\s+(\w+)\s+TO:', joined)
    ship_name = f"{m.group(2)} {m.group(1)}" if m else ""
    m = re.search(r'Date:\s+(\w+)\s+(\w+)', joined)
    ship_loc = f"{m.group(2)} {m.group(1)}" if m else ""
    ship_to = ", ".join(filter(None, [ship_name, ship_loc]))

    m = re.search(r'(\d{1,2}/\d{1,2}/\d{4})', joined)
    date = m.group(1) if m else ""

    qty = ""
    for i, l in enumerate(lines):
        if re.fullmatch(r'\d{1,2}/\d{1,2}/\d{4}', l):
            if i >= 1 and lines[i - 1].isdigit():
                qty = lines[i - 1]
            break

    m = re.search(r'PO#?\s*(\d{6,})', joined)
    po = m.group(1) if m else ""

    m = re.search(r'([A-Z]{2,5}\d{3,}-\d+)', joined)
    style = m.group(1) if m else ""

    gtin = ""
    m = re.search(r'\(01\)(\d+)', joined)
    if m:
        head = m.group(1)
        b = re.search(r'\b(\d{11})\b', joined)
        if b:
            d13 = head + b.group(1)
            gtin = d13 + str(gtin_check_digit(d13))

    color = "Black" if "Black" in joined else ""
    size  = "Prepack" if "Prepack" in joined else ""

    prefix = mid = seq = chk = ""
    for i, l in enumerate(lines):
        if l == "CARTON":
            mid    = lines[i + 1] if i + 1 < len(lines) else ""
            prefix = lines[i + 2] if i + 2 < len(lines) else ""
            seq    = lines[i - 2] if i - 2 >= 0 else ""
            chk    = lines[i - 3] if i - 3 >= 0 else ""
            break
    carton_full = " ".join(filter(None, [prefix, mid, seq, chk]))
    carton_seq  = (mid + seq) if (mid and seq) else seq 
    carton_code = re.sub(r'\D', '', carton_full) 

    desc = ""
    if prefix:
        idxs = [i for i, l in enumerate(lines) if l == prefix]
        if idxs:
            run = []
            for l in lines[idxs[-1] + 1:]:
                if l.startswith("PO#") or (po and l == po):
                    break
                run.append(l)
            desc = " ".join(reversed(run)).strip()

    return {
        "Label No."     : page_no,
        "Ship To"       : ship_to,
        "Date"          : date,
        "PO #"          : po,
        "Style / Color" : style,
        "Description"   : desc,
        "Color"         : color,
        "Size"          : size,
        "Qty"           : qty,
        "GTIN (01)"     : gtin,
        "Carton No."    : carton_full,
        "Carton Seq"    : carton_seq,
        "Carton Barcode": carton_code,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Parser 2 — Bulk / Unichela label format
# ─────────────────────────────────────────────────────────────────────────────
def extract_unichela_labels(text, page_no):
    """
    BULK_1.pdf වැනි එකම පිටුවේ ලේබල් කිහිපයක් තිබිය හැකි format එක සඳහා
    """
    rows = []
    parts = text.split("SHIP TO:")
    
    for part in parts[1:]:
        cleaned = clean_text("SHIP TO: " + part)
        
        date_m = re.search(r'Date:\s*(\d{1,2}/\d{1,2}/\d{4})', cleaned)
        date = date_m.group(1) if date_m else ""
        
        ship_m = re.search(r'Date:.*?\d{4}\s+(.*?)\s+QTY', cleaned)
        ship_to = ship_m.group(1).strip() if ship_m else "MCDONOUGH GA MCDONOUGH DC"
        
        qty_m = re.search(r'QTY\s+(\d+)', cleaned)
        qty = qty_m.group(1) if qty_m else ""
        
        po_m = re.search(r'PO#\s*(\d+)', cleaned)
        po = po_m.group(1) if po_m else ""
        
        # Style / Size හඳුනාගැනීම
        style, size = "", ""
        style_size_m = re.search(r'\b(XXS|XS|S|M|L|XL|XXL|\d+[SML]?)\s+([A-Z0-9]+(?:\s*-\s*[A-Z0-9]+)+)', cleaned)
        if style_size_m:
            size = style_size_m.group(1)
            style = style_size_m.group(2)
            
        # Barcode අංකය හඳුනාගැනීම (දිගු ඉලක්කම් පෙළ)
        barcode = ""
        b_match = re.search(r'QTY\s+\d+\s+([\d\s]{15,})\s+CARTON', cleaned)
        if b_match:
            barcode = b_match.group(1).replace(" ", "")
        else:
            numbers = re.findall(r'\d{12,}', cleaned.replace(" ", ""))
            if numbers:
                barcode = max(numbers, key=len)
                
        rows.append({
            "Label No."     : page_no,
            "Ship To"       : ship_to,
            "Date"          : date,
            "PO #"          : po,
            "Style / Color" : style,
            "Description"   : "UNICHELA / NON-CONFORMING",
            "Color"         : "",
            "Size"          : size,
            "Qty"           : qty,
            "GTIN (01)"     : "",
            "Carton No."    : barcode,
            "Carton Seq"    : "",
            "Carton Barcode": barcode,
        })
    return rows


CARTON_COLUMN_ORDER = [
    "Label No.", "Ship To", "Date", "PO #",
    "Style / Color", "Description", "Color", "Size",
    "Qty", "GTIN (01)", "Carton No.", "Carton Seq", "Carton Barcode",
]

CARTON_COL_WIDTHS = {
    "Label No."     : 9,  "Ship To"       : 28, "Date"          : 12,
    "PO #"          : 14, "Style / Color" : 16, "Description"   : 32,
    "Color"         : 10, "Size"          : 10, "Qty"           : 8,
    "GTIN (01)"     : 18, "Carton No."    : 24, "Carton Seq"    : 13,
    "Carton Barcode": 24,
}


# ─────────────────────────────────────────────────────────────────────────────
# Parser 3 — CMUS / Club Monaco label
# ─────────────────────────────────────────────────────────────────────────────
def extract_label_data(text):
    rows = []

    order_match = re.search(r'Order No\.:(\S+)\s+(\d+)', text)
    order_no = order_match.group(1) if order_match else ""
    seq_no   = order_match.group(2) if order_match else ""

    dest_match = re.search(r'Factory I/O:\s*\n(\S+)', text)
    destination = dest_match.group(1) if dest_match else ""

    ship_from_id_match = re.search(r'Ship From:\s*(\S+)', text)
    ship_from_id = ship_from_id_match.group(1) if ship_from_id_match else ""
    address_block_match = re.search(r'Ship From:.*?\n(.*?)Order No\.', text, re.DOTALL)
    if address_block_match:
        addr_lines = address_block_match.group(1).strip().splitlines()
        ship_from_addr = ", ".join(l.strip() for l in addr_lines if l.strip())
    else:
        ship_from_addr = ""
    ship_from = f"{ship_from_id} – {ship_from_addr}" if ship_from_id else ship_from_addr

    ship_to_header = re.search(r'Ship To:(.+)', text)
    to_name = clean_text(ship_to_header.group(1)) if ship_to_header else ""
    to_addr_match = re.search(r'Export Processing Zone\s+([\w\d].*?)\n.*?Bethlehem', text, re.DOTALL)
    to_street = clean_text(to_addr_match.group(1)) if to_addr_match else ""
    to_city_match = re.search(r'(Bethlehem PA \d+)', text)
    to_city = clean_text(to_city_match.group(1)) if to_city_match else ""
    ship_to = ", ".join(filter(None, [to_name, to_street, to_city]))

    carton_match = re.search(r'CARTON NUMBER\s+(\d+)\s+of\s+(\d+)', text, re.IGNORECASE)
    carton_no     = carton_match.group(1) if carton_match else ""
    total_cartons = carton_match.group(2) if carton_match else ""
    carton_label  = f"{carton_no} of {total_cartons}" if carton_no else ""

    carton_total_match = re.search(r'CARTON TOTAL\s+(\d+)', text, re.IGNORECASE)
    carton_total = carton_total_match.group(1) if carton_total_match else ""

    label_total_match = re.search(r'LABEL TOTAL\s+(\d+)', text, re.IGNORECASE)
    label_total = label_total_match.group(1) if label_total_match else ""

    sscc_raw_match = re.search(r'\(00\)([\d\s]+)', text)
    if sscc_raw_match:
        raw          = sscc_raw_match.group(0).split('\n')[0].strip()
        sscc_display = raw
        sscc_digits  = re.sub(r'[^\d]', '', raw)[:20]
    else:
        sscc_display = sscc_digits = ""

    table_match = re.search(r'Material\s*#\s+Size\s+Quantity\s*\n(.*?)LABEL TOTAL', text, re.DOTALL | re.IGNORECASE)
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
    def _merge(group):
        first = group.iloc[0].copy()
        if len(group) > 1:
            first["Size"]        = "/".join(group["Size"].tolist())
            first["Size Detail"] = "/".join(f'{r["Size"]}{r["Quantity"]}' for _, r in group.iterrows())
            first["Quantity"]    = str(sum(int(q) for q in group["Quantity"] if str(q).isdigit()))
        else:
            first["Size Detail"] = ""
        return first

    return df.groupby("SSCC (digits)", sort=False, group_keys=False).apply(_merge).reset_index(drop=True)


CMUS_COLUMN_ORDER = [
    "Order No.", "Seq No.", "Destination",
    "Ship From", "Ship To",
    "Material #", "Size", "Size Detail", "Quantity",
    "Label Total", "Carton Total", "Carton No.",
    "SSCC (display)", "SSCC (digits)",
]

CMUS_COL_WIDTHS = {
    "Order No."      : 16, "Seq No."   : 8,  "Destination"  : 13,
    "Ship From"      : 40, "Ship To"   : 35, "Material #"   : 18,
    "Size"           : 10, "Size Detail": 22, "Quantity"    : 10,
    "Label Total"    : 12, "Carton Total": 12, "Carton No." : 13,
    "SSCC (display)" : 30, "SSCC (digits)": 25,
}


# ─────────────────────────────────────────────────────────────────────────────
# Excel builder (shared)
# ─────────────────────────────────────────────────────────────────────────────
def build_excel(df, col_widths, summary_df=None, highlight_mixed=True):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Shipping_Data')
        workbook  = writer.book
        worksheet = writer.sheets['Shipping_Data']

        header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#1E1E1E', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter',
        })
        data_fmt  = workbook.add_format({'border': 1, 'valign': 'vcenter'})
        mixed_fmt = workbook.add_format({
            'border': 1, 'valign': 'vcenter', 'bg_color': '#FFF9C4',
        })

        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, header_fmt)
            worksheet.set_column(col_num, col_num, col_widths.get(col_name, 18))

        for row_idx, row in df.iterrows():
            is_mixed = highlight_mixed and "/" in str(row.get("Size", ""))
            fmt = mixed_fmt if is_mixed else data_fmt
            for col_idx, col_name in enumerate(df.columns):
                worksheet.write(row_idx + 1, col_idx, row[col_name], fmt)

        worksheet.set_row(0, 20)
        worksheet.freeze_panes(1, 0)

        if summary_df is not None and not summary_df.empty:
            summary_df.to_excel(writer, index=False, sheet_name='Summary')
            ws2 = writer.sheets['Summary']
            for col_num, col_name in enumerate(summary_df.columns):
                ws2.write(0, col_num, col_name, header_fmt)
                ws2.set_column(col_num, col_num, max(14, len(str(col_name)) + 4))
            ws2.set_row(0, 20)

    return output.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────
if uploaded_file is not None:
    try:
        with st.spinner("දත්ත කියවමින් පවතී..."):
            carton_rows = []
            cmus_rows   = []
            with pdfplumber.open(uploaded_file) as pdf:
                total_pages = len(pdf.pages)
                progress = st.progress(0.0)
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text()
                    fmt  = detect_format(text)
                    
                    if fmt == "carton":
                        carton_rows.append(parse_carton_label(text, i + 1))
                    elif fmt == "unichela":
                        carton_rows.extend(extract_unichela_labels(text, i + 1))
                    elif fmt == "cmus":
                        cmus_rows.extend(extract_label_data(text))
                        
                    if total_pages:
                        progress.progress((i + 1) / total_pages)
                progress.empty()

        # ── Carton & Unichela Formats ───────────────────────────────────────
        if carton_rows:
            df = pd.DataFrame(carton_rows)
            for col in CARTON_COLUMN_ORDER:
                if col not in df.columns:
                    df[col] = ""
            df = df[CARTON_COLUMN_ORDER].astype(str).replace("nan", "")
            
            # Barcode duplicate වීම වැළැක්වීම
            df = df.drop_duplicates(subset=['Carton Barcode'], keep='first')

            total_cartons = len(df)
            total_qty = sum(int(q) for q in df["Qty"] if str(q).isdigit())

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("මුළු Cartons (Unique)", f"{total_cartons:,}")
            c2.metric("මුළු Qty", f"{total_qty:,}")
            c3.metric("PO #", df["PO #"].iloc[0] if total_cartons else "-")
            c4.metric("Style", df["Style / Color"].iloc[0] if total_cartons else "-")

            st.success(f"✅ Cartons {total_cartons:,} ක දත්ත සාර්ථකව හඳුනා ගන්නා ලදී! (Duplicates ඉවත් කරන ලදී)")
            st.dataframe(df, use_container_width=True, height=420)

            summary = (
                df.assign(_q=pd.to_numeric(df["Qty"], errors="coerce").fillna(0).astype(int))
                  .groupby(["Ship To", "PO #", "Style / Color", "Color", "Size"], as_index=False)
                  .agg(**{"Cartons": ("_q", "size"), "Total Qty": ("_q", "sum")})
            )
            summary["Total Qty"] = summary["Total Qty"].astype(str)
            summary["Cartons"]   = summary["Cartons"].astype(str)

            with st.expander("📋 සාරාංශය (Summary)"):
                st.dataframe(summary, use_container_width=True)

            excel_bytes = build_excel(df, CARTON_COL_WIDTHS, summary_df=summary, highlight_mixed=False)
            st.download_button(
                label="📥 Download Master Excel File",
                data=excel_bytes,
                file_name="Carton_Master_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        # ── CMUS format ────────────────────────────────────────────────────
        elif cmus_rows:
            df = pd.DataFrame(cmus_rows)
            df = merge_sscc_groups(df)
            for col in CMUS_COLUMN_ORDER:
                if col not in df.columns:
                    df[col] = ""
            df = df[CMUS_COLUMN_ORDER].astype(str).replace("nan", "")
            
            # CMUS format එක සඳහා Barcode/SSCC duplicate වීම වැළැක්වීම
            df = df.drop_duplicates(subset=['SSCC (digits)'], keep='first')

            total_labels = len(df)
            st.success(f"✅ ලේබල් {total_labels} ක දත්ත සාර්ථකව හඳුනා ගන්නා ලදී! (Duplicates ඉවත් කරන ලදී)")
            st.dataframe(df, use_container_width=True)

            excel_bytes = build_excel(df, CMUS_COL_WIDTHS, highlight_mixed=True)
            st.download_button(
                label="📥 Download Master Excel File",
                data=excel_bytes,
                file_name="Shipping_Master_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        else:
            st.warning("PDF එකෙන් දත්ත හඳුනා ගත නොහැකි විය. (Format එක support නොකරයි)")

    except Exception as e:
        st.error(f"දෝෂයක් සිදුවිය: {e}")
        raise

st.markdown("---")
st.caption("© 2024 AI Shipping Tool | Developed by **Lakshan**")
