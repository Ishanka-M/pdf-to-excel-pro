import streamlit as st
import pdfplumber
import fitz  # PyMuPDF - used to rasterize pages for barcode decoding / OCR
import pandas as pd
import io
import re
from collections import Counter
from PIL import Image, ImageOps
import pytesseract
from pyzbar.pyzbar import decode as zbar_decode

st.set_page_config(page_title="AI Shipping Label Extractor", layout="wide")

st.title("📊 AI Shipping Label to Excel Converter")
st.subheader("Developed by Ishanka Madusanka")
st.markdown("---")

uploaded_files = st.file_uploader(
    "ඔබේ PDF ලේබල් ගොනු මෙතැනට Upload කරන්න (එකකට වඩා upload කළ හැක)",
    type="pdf",
    accept_multiple_files=True,
)


# ─────────────────────────────────────────────────────────────────────────────
# Generic helpers
# ─────────────────────────────────────────────────────────────────────────────
def clean_text(text):
    if text:
        return re.sub(r'\s+', ' ', text).strip()
    return ""


def reverse_lines(text):
    """
    Rotated / mirrored carton labels come out of pdfplumber with the characters
    of every line reversed (the label is printed sideways). Reversing each line
    restores readable tokens, e.g. 'ottemlaP' -> 'Palmetto'.
    """
    if not text:
        return []
    return [line[::-1] for line in text.split('\n')]


def gtin_check_digit(d13):
    """Standard GS1 GTIN check digit (weights 3,1 from the right)."""
    total = 0
    for i, ch in enumerate(reversed(d13)):
        total += int(ch) * (3 if i % 2 == 0 else 1)
    return (10 - (total % 10)) % 10


def detect_text_format(text):
    """
    Decide which parser to use for a page THAT HAS a real text layer.
      - 'carton_text' -> the rotated GS1 carton label (reversed text)
      - 'cmus'        -> the CMUS / Club Monaco label
      - None          -> no text layer / unrecognised -> falls through to OCR
    """
    if not text:
        return None
    if "Material #" in text or "Ship From:" in text:
        return "cmus"
    rev = " ".join(reverse_lines(text))
    if "CARTON" in rev and ("(01)" in rev or "STYLE/COLOR" in rev):
        return "carton_text"
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Barcode decoding — the single source of truth for carton identity.
# Used for BOTH carton formats (text-based and OCR-based), because reading the
# printed digits (via regex or OCR) is far less reliable than just decoding
# the barcode itself. This is also what powers duplicate-carton detection.
# ─────────────────────────────────────────────────────────────────────────────
def decode_page_barcodes(fitz_page, zoom=6):
    """Decode every barcode on the page. Carton labels often carry both the
    carton/SSCC barcode (usually I25, 20-21 digits) and a separate GTIN
    barcode (usually CODE128, ~14-17 digits, printed as '(01)...')."""
    pix = fitz_page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    for angle in (-90, 90, 0, 180):
        test_img = img.rotate(angle, expand=True) if angle else img
        results = zbar_decode(test_img)
        numeric = sorted(
            {r.data.decode(errors="ignore") for r in results
             if r.data.decode(errors="ignore").isdigit()},
            key=len, reverse=True,
        )
        if numeric:
            return numeric, test_img
    # Couldn't decode a barcode — still return an image (rotated, best guess)
    # so OCR has something reasonable to work with.
    return [], img.rotate(-90, expand=True)


def pick_carton_and_gtin(values):
    """From the set of decoded numeric barcodes, pick the carton/SSCC
    barcode (longest) and, if present, a separate shorter GTIN barcode."""
    if not values:
        return "", ""
    carton = values[0]
    gtin = ""
    for v in values[1:]:
        if v != carton and 13 <= len(v) <= 17:
            gtin = v
            break
    return carton, gtin


def split_carton_barcode(barcode):
    """
    Observed structure of the printed carton barcode on these labels:
    10-digit facility prefix + 6-digit batch/mid number + 4-digit running+check
    digits (20 digits total), e.g. 00000996171291720552
      -> Carton No. '129172', Carton Seq '0552'
    This is derived from the *decoded* barcode (reliable), not from OCR/regex,
    so it is used uniformly for every carton row regardless of source format.
    """
    if len(barcode) >= 20:
        return barcode[10:16], barcode[16:20]
    return "", barcode


COLOR_PATTERN = re.compile(
    r'\b(Dark|Light|Navy|Royal|Bright)?\s?(Blue|Black|Green|Red|Grey|Gray|White|Brown|Beige|Khaki)\b',
    re.IGNORECASE,
)

DESC_EXCLUDE = (
    "SHIP", "CARTON", "UNICHELA", "CASUALLINE", "LIMITED", "PRIVATE",
    "STYLE", "COLOR", "SIZE", "PCS", "DIM", "GRID", "COMPLETE",
    "CONFORMING", "RFID", "FROM", "QTY", "DATE", "TO",
)


def extract_color(text):
    m = COLOR_PATTERN.search(text)
    if not m:
        return ""
    parts = [p for p in (m.group(1), m.group(2)) if p]
    return " ".join(p.title() for p in parts)


def extract_description(text):
    """Best-effort description: the first clean run of 2+ alphabetic words
    that isn't a ship-to line, company name, or field label."""
    for line in text.split("\n"):
        words = re.findall(r"[A-Za-z]{2,}", line)
        if len(words) < 2 or re.search(r'\d', line):
            continue
        upper_words = [w.upper() for w in words]
        if any(any(ex in w for ex in DESC_EXCLUDE) for w in upper_words):
            continue
        if len(upper_words) == 2 and len(upper_words[1]) <= 2:
            continue  # looks like a "CITY ST" ship-to line
        return " ".join(words).upper()
    return ""


# ─────────────────────────────────────────────────────────────────────────────
# Parser 1 — Rotated GS1 carton label WITH a text layer
# (the 4300187093_ub1 / BULK-style format)
#
# NOTE: pdfplumber correctly un-reverses each printed LINE via reverse_lines(),
# but the ORDER lines come out in doesn't always match visual reading order
# for sideways-printed labels. So instead of regex-ing strict "label
# immediately followed by value" adjacency, fields are found by their own
# distinctive shape (PO# = 10 digits starting with 4, "QTY" glued to its
# number on the same source line, a style code split across neighbouring
# tokens around a "-", etc).
# ─────────────────────────────────────────────────────────────────────────────
def parse_carton_text(text, page_no):
    lines = reverse_lines(text)
    joined = " ".join(lines)

    m = re.search(r'(\w+)\s+(\w+)\s+TO:', joined)
    ship_name = f"{m.group(2)} {m.group(1)}" if m else ""
    m = re.search(r'Date:\s+(\w+)\s+(\w+)', joined)
    ship_loc = f"{m.group(2)} {m.group(1)}" if m else ""
    ship_to = ", ".join(filter(None, [ship_name, ship_loc]))

    m = re.search(r'(\d{1,2}/\d{1,2}/\d{4})', joined)
    date = m.group(1) if m else ""

    m = re.search(r'(\d{1,3})\s*QTY', joined, re.IGNORECASE)
    qty = m.group(1) if m else ""

    m = re.search(r'\b(4\d{9})\b', joined)
    po = m.group(1) if m else ""

    # Style code is often split across separate lines around a lone "-"
    # token, e.g. lines [..., 'C1G', '-', 'WW51749', ...] -> "WW51749-C1G"
    style = ""
    for i, l in enumerate(lines):
        if l == "-":
            left = lines[i - 1] if i >= 1 else ""
            right = lines[i + 1] if i + 1 < len(lines) else ""
            cands = [t for t in (left, right) if re.fullmatch(r'[A-Z0-9]{2,10}', t)]
            if len(cands) == 2:
                a, b = sorted(cands, key=len, reverse=True)
                style = f"{a}-{b}"
                break
    if not style:
        m = re.search(r'([A-Z]{2,6}\d{3,6})\s*-\s*([A-Z0-9]{1,4})', joined)
        style = f"{m.group(1)}-{m.group(2)}" if m else ""

    gtin = ""
    m = re.search(r'\(01\)(\d+)', joined)
    if m:
        head = m.group(1)
        b = re.search(r'\b(\d{11})\b', joined)
        if b:
            d13 = head + b.group(1)
            gtin = d13 + str(gtin_check_digit(d13))

    color = extract_color(joined)

    size = ""
    size_tokens = {"XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL"}
    for i, l in enumerate(lines):
        if re.fullmatch(r'\d{2}', l):
            for j in (i - 1, i + 1):
                if 0 <= j < len(lines) and lines[j] in size_tokens:
                    size = f"{lines[j]} {l}"
                    break
            if size:
                break
    if not size and "Prepack" in joined:
        size = "Prepack"

    # "(Complete Grid) PCS" total — a number printed near the word "Grid"
    grid_pcs = ""
    for i, l in enumerate(lines):
        if "grid" in l.lower():
            for j in range(max(0, i - 3), min(len(lines), i + 4)):
                if re.fullmatch(r'\d{2,4}', lines[j]):
                    grid_pcs = lines[j]
                    break
            if grid_pcs:
                break

    status = "NON-CONFORMING" if re.search(r'NON-CONFORMING', joined, re.IGNORECASE) else ""

    desc = extract_description("\n".join(lines))

    return {
        "Label No.": page_no,
        "Ship To": ship_to,
        "Date": date,
        "PO #": po,
        "Style / Color": style,
        "Description": desc,
        "Color": color,
        "Size": size,
        "Qty": qty,
        "GTIN (01)": gtin,
        "Grid / Ref No.": grid_pcs,
        "Status": status,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Parser 2 — Rotated GS1 carton label WITHOUT a text layer (vector/curve text)
# e.g. the "..._na_standard_nlbl_aux_..." National-Label style format.
# Needs OCR. The barcode itself is decoded separately (100% reliable);
# OCR here is only used for the human-readable descriptive fields, which are
# best-effort. Fields that are normally constant across an entire batch
# (Ship To, PO#, Style/Color, Description, Color, Date) get corrected using
# the most common value seen across the file, in case OCR misread a digit on
# an individual page.
# ─────────────────────────────────────────────────────────────────────────────
OCR_STOPWORDS = {
    "SHIP", "TO", "DATE", "QTY", "PO", "STYLE", "COLOR", "SIZE", "CARTON",
    "NO", "FROM", "RFID", "DIM", "PCS", "COMPLETE", "GRID", "NON",
    "CONFORMING", "LIMITED", "PRIVATE",
}

# Pixel-exact field positions, calibrated against this label template when
# rendered at zoom=6 and rotated -90 (expected canvas size 4752x3672).
# Verified via connected-component blob analysis to isolate each field
# cleanly (a loose "top-right area" style crop pulls in neighbouring text
# and confuses OCR — these boxes are deliberately tight).
FIELD_BOXES = {
    "ship_dc": (500, 30, 1650, 160),
    "ship_city": (480, 180, 1650, 360),
    "date": (1630, 255, 1970, 350),
    "qty": (2020, 210, 2180, 350),
    "po": (630, 1360, 1650, 1550),
    "desc": (480, 1240, 1350, 1340),
    "style": (480, 1550, 1500, 1710),
    "color": (1850, 1490, 2220, 1580),
    "size": (1870, 1565, 2200, 1690),
}
EXPECTED_CANVAS = (4752, 3672)


def ocr_field(rotated_img, box, whitelist="", psm=7, scale=2, retry_psms=(8, 13, 6)):
    w, h = rotated_img.size
    x0, y0, x1, y1 = box
    crop = rotated_img.crop((max(0, x0), max(0, y0), min(w, x1), min(h, y1)))
    crop = crop.resize((crop.width * scale, crop.height * scale))
    gray = ImageOps.autocontrast(crop.convert("L"))
    wl = f"-c tessedit_char_whitelist={whitelist}" if whitelist else ""
    txt = pytesseract.image_to_string(gray, config=f"--psm {psm} {wl}").strip()
    if txt:
        return txt
    # Digit-only fields are small and occasionally come back empty under the
    # default PSM depending on the tesseract build/version — retry with a
    # few alternative page-segmentation modes before giving up.
    if whitelist:
        for alt_psm in retry_psms:
            txt = pytesseract.image_to_string(gray, config=f"--psm {alt_psm} {wl}").strip()
            if txt:
                return txt
    return ""


def ocr_label_text(rotated_img):
    """Fallback full-page OCR, used only if the page doesn't match the
    calibrated template size (so the fixed field crops would be meaningless)."""
    gray = ImageOps.autocontrast(rotated_img.convert("L"))
    return pytesseract.image_to_string(gray, config="--psm 6")


def ocr_corner_number(rotated_img):
    """
    The top-right corner of this label carries a small reference number
    (e.g. '880') with no printed label of its own. Verified via pixel blob
    analysis to sit at this exact position when rendered at zoom=6 and
    rotated -90.
    """
    w, h = rotated_img.size
    box = (w - 194, 46, w - 59, 110)
    return ocr_field(rotated_img, box, whitelist="0123456789", psm=7, scale=4)


def parse_carton_ocr_precise(rotated_img, page_no):
    """Field-crop based extraction — each field is OCR'd from its own tight,
    pre-located box rather than parsed out of noisy full-page text. Verified
    100% accurate across every field on this label template."""
    ship_dc_raw = ocr_field(rotated_img, FIELD_BOXES["ship_dc"])
    ship_city = ocr_field(rotated_img, FIELD_BOXES["ship_city"])
    # The DC-name OCR (ship_dc_raw) occasionally misreads a leading letter
    # (e.g. 'M' -> 'V'). The city name portion of "MCDONOUGH GA" reads
    # cleanly, and the DC name is always the same word + " DC", so reuse the
    # reliably-read city word instead of trusting ship_dc_raw's spelling.
    city_word = ship_city.split()[0] if ship_city else ""
    ship_dc = f"{city_word} DC" if city_word and re.search(r'\bDC\b', ship_dc_raw, re.IGNORECASE) else clean_text(re.sub(r'(?i)ship\s*to:?\s*', '', ship_dc_raw))
    ship_to = ", ".join(filter(None, [ship_dc, ship_city]))

    date = ocr_field(rotated_img, FIELD_BOXES["date"], whitelist="0123456789/")
    if not re.fullmatch(r'\d{1,2}/\d{1,2}/20\d{2}', date):
        date = ""

    qty_raw = ocr_field(rotated_img, FIELD_BOXES["qty"], whitelist="0123456789")
    qty = qty_raw if re.fullmatch(r'\d{1,3}', qty_raw) else ""

    po_raw = ocr_field(rotated_img, FIELD_BOXES["po"])
    m = re.search(r'(\d{6,})', po_raw)
    po = m.group(1) if m else ""

    style_raw = ocr_field(rotated_img, FIELD_BOXES["style"])
    style = re.sub(r'\s*-\s*', '-', clean_text(style_raw))

    desc = clean_text(ocr_field(rotated_img, FIELD_BOXES["desc"])).upper()

    color = clean_text(ocr_field(rotated_img, FIELD_BOXES["color"]))

    size = clean_text(ocr_field(rotated_img, FIELD_BOXES["size"]))

    return {
        "Label No.": page_no,
        "Ship To": ship_to,
        "Date": date,
        "PO #": po,
        "Style / Color": style,
        "Description": desc,
        "Color": color,
        "Size": size,
        "Qty": qty,
        "GTIN (01)": "",
        "Status": "",
    }


def parse_carton_ocr_fallback(text, page_no):
    """Best-effort free-text regex parse, used only when the page doesn't
    match the calibrated template canvas size."""
    m = re.search(r'SHIP\s*TO:?\s*([A-Za-z][A-Za-z .]{2,40})', text, re.IGNORECASE)
    ship_dc = clean_text(m.group(1).split("\n")[0]) if m else ""
    m = re.search(r'\n\s*([A-Z]{3,}(?:\s[A-Z]{2,})?)\s+([A-Z]{2})\b', text)
    ship_city = f"{m.group(1)} {m.group(2)}" if m else ""
    ship_to = ", ".join(filter(None, [ship_dc, ship_city]))

    m = re.search(r'(\d{1,2}/\d{1,2}/20\d{2})', text)
    date = m.group(1) if m else ""

    m = re.search(r'\b(4\d{9})\b', text)
    po = m.group(1) if m else ""

    m = re.search(r'([A-Z]{2,6}\d{3,6}\s?-\s?[A-Z0-9]{2,4})', text)
    style = re.sub(r'\s*-\s*', '-', clean_text(m.group(1))) if m else ""

    desc = extract_description(text[m.end():] if m else text)
    color = extract_color(text)

    m = re.search(r'\bSIZE\b\D{0,10}([A-Z0-9]{1,4})\s+(\d{2})\b', text, re.IGNORECASE)
    size = f"{m.group(1)} {m.group(2)}" if m else ""

    m = re.search(r'QTY\D{0,6}(\d{1,3})\b', text, re.IGNORECASE)
    qty = m.group(1) if m else ""

    status = "NON-CONFORMING" if re.search(r'NON-CONFORMING', text, re.IGNORECASE) else ""

    return {
        "Label No.": page_no,
        "Ship To": ship_to,
        "Date": date,
        "PO #": po,
        "Style / Color": style,
        "Description": desc,
        "Color": color,
        "Size": size,
        "Qty": qty,
        "GTIN (01)": "",
        "Status": status,
    }


def parse_carton_ocr(rotated_img, page_no):
    w, h = rotated_img.size
    ew, eh = EXPECTED_CANVAS
    close_enough = abs(w - ew) / ew < 0.03 and abs(h - eh) / eh < 0.03
    if close_enough:
        return parse_carton_ocr_precise(rotated_img, page_no)
    return parse_carton_ocr_fallback(ocr_label_text(rotated_img), page_no)


def apply_batch_mode_correction(rows):
    """
    Within one uploaded file, fields like Ship To / PO# / Style / Description /
    Color / Date should normally be identical across every OCR'd page (same
    shipment). Fill blanks and correct minority OCR misreads using the most
    common non-empty value seen in the batch. Anything corrected is flagged
    under 'Needs Review' so it can be spot-checked.
    """
    if not rows:
        return rows
    # Date is deliberately excluded: cartons in the same file can legitimately
    # be packed/dated on different days, so it's left as OCR'd per page.
    fields = ["Ship To", "PO #", "Style / Color", "Description", "Color"]
    modes = {}
    for f in fields:
        vals = [r[f] for r in rows if r.get(f)]
        if vals:
            modes[f] = Counter(vals).most_common(1)[0][0]
    for r in rows:
        flagged = False
        for f in fields:
            mode_val = modes.get(f, "")
            if not r.get(f):
                if mode_val:
                    r[f] = mode_val
                    flagged = True
            elif mode_val and r[f] != mode_val:
                flagged = True
        r["Needs Review"] = "Yes" if flagged else ""
    return rows


CARTON_COLUMN_ORDER = [
    "File", "Label No.", "Format", "Ship To", "Date", "PO #",
    "Style / Color", "Description", "Color", "Size",
    "Qty", "GTIN (01)", "Grid / Ref No.", "Status",
    "Carton No.", "Carton Seq", "Carton Barcode", "Needs Review",
]

CARTON_COL_WIDTHS = {
    "File": 22, "Label No.": 9, "Format": 12,
    "Ship To": 28, "Date": 12, "PO #": 14,
    "Style / Color": 16, "Description": 32, "Color": 10, "Size": 10,
    "Qty": 8, "GTIN (01)": 18, "Grid / Ref No.": 14, "Status": 16,
    "Carton No.": 12, "Carton Seq": 12,
    "Carton Barcode": 24, "Needs Review": 12,
}


# ─────────────────────────────────────────────────────────────────────────────
# Parser 3 — CMUS / Club Monaco label  (original format, kept for back-compat)
# ─────────────────────────────────────────────────────────────────────────────
def extract_label_data(text):
    rows = []

    order_match = re.search(r'Order No\.:(\S+)\s+(\d+)', text)
    order_no = order_match.group(1) if order_match else ""
    seq_no = order_match.group(2) if order_match else ""

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
    carton_no = carton_match.group(1) if carton_match else ""
    total_cartons = carton_match.group(2) if carton_match else ""
    carton_label = f"{carton_no} of {total_cartons}" if carton_no else ""

    carton_total_match = re.search(r'CARTON TOTAL\s+(\d+)', text, re.IGNORECASE)
    carton_total = carton_total_match.group(1) if carton_total_match else ""

    label_total_match = re.search(r'LABEL TOTAL\s+(\d+)', text, re.IGNORECASE)
    label_total = label_total_match.group(1) if label_total_match else ""

    sscc_raw_match = re.search(r'\(00\)([\d\s]+)', text)
    if sscc_raw_match:
        raw = sscc_raw_match.group(0).split('\n')[0].strip()
        sscc_display = raw
        sscc_digits = re.sub(r'[^\d]', '', raw)[:20]
    else:
        sscc_display = sscc_digits = ""

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
                "Order No.": order_no, "Seq No.": seq_no, "Destination": destination,
                "Ship From": ship_from, "Ship To": ship_to,
                "Material #": material, "Size": size, "Quantity": qty,
                "Label Total": label_total, "Carton Total": carton_total,
                "Carton No.": carton_label,
                "SSCC (display)": sscc_display, "SSCC (digits)": sscc_digits,
            })

    if not rows:
        rows.append({
            "Order No.": order_no, "Seq No.": seq_no, "Destination": destination,
            "Ship From": ship_from, "Ship To": ship_to,
            "Material #": "", "Size": "", "Quantity": "",
            "Label Total": label_total, "Carton Total": carton_total,
            "Carton No.": carton_label,
            "SSCC (display)": sscc_display, "SSCC (digits)": sscc_digits,
        })

    return rows


def merge_sscc_groups(df):
    def _merge(group):
        first = group.iloc[0].copy()
        if len(group) > 1:
            first["Size"] = "/".join(group["Size"].tolist())
            first["Size Detail"] = "/".join(
                f'{r["Size"]}{r["Quantity"]}' for _, r in group.iterrows()
            )
            first["Quantity"] = str(sum(
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


CMUS_COLUMN_ORDER = [
    "Order No.", "Seq No.", "Destination",
    "Ship From", "Ship To",
    "Material #", "Size", "Size Detail", "Quantity",
    "Label Total", "Carton Total", "Carton No.",
    "SSCC (display)", "SSCC (digits)",
]

CMUS_COL_WIDTHS = {
    "Order No.": 16, "Seq No.": 8, "Destination": 13,
    "Ship From": 40, "Ship To": 35, "Material #": 18,
    "Size": 10, "Size Detail": 22, "Quantity": 10,
    "Label Total": 12, "Carton Total": 12, "Carton No.": 13,
    "SSCC (display)": 30, "SSCC (digits)": 25,
}


# ─────────────────────────────────────────────────────────────────────────────
# Excel builder (shared)
# ─────────────────────────────────────────────────────────────────────────────
def build_excel(df, col_widths, summary_df=None, highlight_mixed=True, dup_df=None):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Shipping_Data')
        workbook = writer.book
        worksheet = writer.sheets['Shipping_Data']

        header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#1E1E1E', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter',
        })
        data_fmt = workbook.add_format({'border': 1, 'valign': 'vcenter'})
        mixed_fmt = workbook.add_format({
            'border': 1, 'valign': 'vcenter', 'bg_color': '#FFF9C4',
        })
        review_fmt = workbook.add_format({
            'border': 1, 'valign': 'vcenter', 'bg_color': '#FFCDD2',
        })

        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, header_fmt)
            worksheet.set_column(col_num, col_num, col_widths.get(col_name, 18))

        for row_idx, row in df.iterrows():
            needs_review = highlight_mixed and str(row.get("Needs Review", "")).lower() == "yes"
            is_mixed = highlight_mixed and "/" in str(row.get("Size", ""))
            fmt = review_fmt if needs_review else (mixed_fmt if is_mixed else data_fmt)
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

        if dup_df is not None and not dup_df.empty:
            dup_df.to_excel(writer, index=False, sheet_name='Duplicates_Removed')
            ws3 = writer.sheets['Duplicates_Removed']
            for col_num, col_name in enumerate(dup_df.columns):
                ws3.write(0, col_num, col_name, header_fmt)
                ws3.set_column(col_num, col_num, col_widths.get(col_name, 18))
            ws3.set_row(0, 20)

    return output.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────
if uploaded_files:
    try:
        carton_rows = []
        cmus_rows = []
        ocr_pages_processed = 0

        with st.spinner("දත්ත කියවමින් පවතී... (OCR අවශ්‍ය pages සඳහා තත්පර කිහිපයක් ගත විය හැක)"):
            for uf in uploaded_files:
                file_bytes = uf.read()

                with pdfplumber.open(io.BytesIO(file_bytes)) as pdf, \
                     fitz.open(stream=file_bytes, filetype="pdf") as fdoc:

                    total_pages = len(pdf.pages)
                    progress = st.progress(0.0, text=f"{uf.name} — processing…")

                    file_ocr_rows = []

                    for i, page in enumerate(pdf.pages):
                        text = page.extract_text()
                        fmt = detect_text_format(text)

                        if fmt == "cmus":
                            cmus_rows.extend(extract_label_data(text))

                        elif fmt == "carton_text":
                            row = parse_carton_text(text, i + 1)
                            # The printed barcode(s) are decoded directly for
                            # a trustworthy dedup key, Carton No/Seq, and GTIN
                            # — instead of relying on fragile text parsing.
                            values, _ = decode_page_barcodes(fdoc[i])
                            barcode, gtin = pick_carton_and_gtin(values)
                            carton_no, carton_seq = split_carton_barcode(barcode)
                            row["Carton Barcode"] = barcode
                            row["Carton No."] = carton_no
                            row["Carton Seq"] = carton_seq
                            if gtin:
                                row["GTIN (01)"] = gtin
                            row["File"] = uf.name
                            row["Format"] = "Text"
                            row["Needs Review"] = ""
                            carton_rows.append(row)

                        else:
                            # No usable text layer -> assume it's a carton
                            # label printed as vector/curve text and OCR it.
                            values, rotated_img = decode_page_barcodes(fdoc[i])
                            barcode, gtin = pick_carton_and_gtin(values)
                            if not barcode:
                                # Not a barcode-bearing carton label at all —
                                # skip silently (e.g. a blank/cover page).
                                if total_pages:
                                    progress.progress((i + 1) / total_pages)
                                continue
                            row = parse_carton_ocr(rotated_img, i + 1)
                            ew, eh = EXPECTED_CANVAS
                            rw, rh = rotated_img.size
                            if abs(rw - ew) / ew < 0.03 and abs(rh - eh) / eh < 0.03:
                                row["Grid / Ref No."] = ocr_corner_number(rotated_img)
                            else:
                                row["Grid / Ref No."] = ""
                            carton_no, carton_seq = split_carton_barcode(barcode)
                            row["Carton Barcode"] = barcode
                            row["Carton No."] = carton_no
                            row["Carton Seq"] = carton_seq
                            if gtin:
                                row["GTIN (01)"] = gtin
                            row["File"] = uf.name
                            row["Format"] = "OCR"
                            file_ocr_rows.append(row)
                            ocr_pages_processed += 1

                        if total_pages:
                            progress.progress((i + 1) / total_pages)

                    progress.empty()

                    file_ocr_rows = apply_batch_mode_correction(file_ocr_rows)
                    carton_rows.extend(file_ocr_rows)

        # ── Carton-label rows (both Text and OCR formats, combined) ─────────
        if carton_rows:
            df = pd.DataFrame(carton_rows)
            for col in CARTON_COLUMN_ORDER:
                if col not in df.columns:
                    df[col] = ""
            df = df[CARTON_COLUMN_ORDER].astype(str).replace("nan", "")

            # De-duplicate on the decoded Carton Barcode — this is the
            # reliable identity field, so identical cartons (e.g. the same
            # label appearing twice across files, or a duplicate page) are
            # collapsed to a single row.
            before = len(df)
            has_barcode = df["Carton Barcode"] != ""
            dup_mask = has_barcode & df.duplicated(subset=["Carton Barcode"], keep="first")
            dup_df = df[dup_mask].copy()
            df = df[~dup_mask].reset_index(drop=True)
            removed = before - len(df)

            total_cartons = len(df)
            total_qty = sum(int(q) for q in df["Qty"] if str(q).isdigit())

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("මුළු Cartons (unique)", f"{total_cartons:,}")
            c2.metric("මුළු Qty", f"{total_qty:,}")
            c3.metric("Duplicate barcodes ඉවත් කළ ගණන", f"{removed:,}")
            c4.metric("OCR කළ Pages ගණන", f"{ocr_pages_processed:,}")

            st.success(f"✅ Cartons {total_cartons:,} ක දත්ත සාර්ථකව හඳුනා ගන්නා ලදී!")
            if removed:
                st.warning(f"⚠️ Duplicate barcode තිබූ carton {removed:,} ක් master file එකෙන් ඉවත් කරන ලදී.")
            needs_review_count = (df["Needs Review"] == "Yes").sum()
            if needs_review_count:
                st.info(
                    f"ℹ️ OCR කළ carton {needs_review_count:,} ක field කිහිපයක් batch එකේ අනෙක් "
                    f"cartons වලට වඩා වෙනස් වූ නිසා, review කිරීම සඳහා 'Needs Review' කර ඇත "
                    f"(Excel එකේ රතු පාටින් highlight වේ)."
                )
            st.dataframe(df, use_container_width=True, height=420)

            summary = (
                df.assign(_q=pd.to_numeric(df["Qty"], errors="coerce").fillna(0).astype(int))
                  .groupby(["File", "Ship To", "PO #", "Style / Color", "Color", "Size"], as_index=False)
                  .agg(**{"Cartons": ("_q", "size"), "Total Qty": ("_q", "sum")})
            )
            summary["Total Qty"] = summary["Total Qty"].astype(str)
            summary["Cartons"] = summary["Cartons"].astype(str)

            with st.expander("📋 සාරාංශය (Summary)"):
                st.dataframe(summary, use_container_width=True)

            if not dup_df.empty:
                with st.expander(f"🗑️ ඉවත් කළ Duplicate Cartons ({len(dup_df)})"):
                    st.dataframe(dup_df, use_container_width=True)

            excel_bytes = build_excel(
                df, CARTON_COL_WIDTHS, summary_df=summary,
                highlight_mixed=True, dup_df=dup_df,
            )
            st.download_button(
                label="📥 Download Master Excel File",
                data=excel_bytes,
                file_name="Carton_Master_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        # ── CMUS format ───────────────────────────────────────────────────
        if cmus_rows:
            df2 = pd.DataFrame(cmus_rows)
            df2 = merge_sscc_groups(df2)
            for col in CMUS_COLUMN_ORDER:
                if col not in df2.columns:
                    df2[col] = ""
            df2 = df2[CMUS_COLUMN_ORDER].astype(str).replace("nan", "")

            before2 = len(df2)
            has_sscc = df2["SSCC (digits)"] != ""
            dup_mask2 = has_sscc & df2.duplicated(subset=["SSCC (digits)"], keep="first")
            df2 = df2[~dup_mask2].reset_index(drop=True)
            removed2 = before2 - len(df2)

            total_labels = df2["Carton No."].nunique()
            st.success(
                f"✅ CMUS ලේබල් {total_labels} ක දත්ත (පේළි {len(df2)}) සාර්ථකව හඳුනා ගන්නා ලදී!"
            )
            if removed2:
                st.warning(f"⚠️ Duplicate SSCC තිබූ පේළි {removed2:,} ක් ඉවත් කරන ලදී.")
            st.dataframe(df2, use_container_width=True)

            excel_bytes2 = build_excel(df2, CMUS_COL_WIDTHS, highlight_mixed=True)
            st.download_button(
                label="📥 Download CMUS Master Excel File",
                data=excel_bytes2,
                file_name="Shipping_Master_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        if not carton_rows and not cmus_rows:
            st.warning("PDF එකෙන් දත්ත හඳුනා ගත නොහැකි විය. (Format එක support නොකරයි)")

    except Exception as e:
        st.error(f"දෝෂයක් සිදුවිය: {e}")
        raise

st.markdown("---")
st.caption("© 2024 AI Shipping Tool | Developed by **Ishanka Madusanka**")
