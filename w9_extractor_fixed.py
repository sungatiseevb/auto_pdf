"""
W-9 PDF Data Extraction Pipeline - FIXED VERSION
Key fixes:
1. Improved business_name extraction from scanned PDFs
2. Better SSN/EIN regex for OCR noise
3. Improved date extraction from signature section
4. Always uses openpyxl (no xlwings issues)
"""

import re
import os
import shutil
import logging
import unicodedata
from pathlib import Path

import warnings
warnings.filterwarnings("ignore")
os.environ["TOKENIZERS_PARALLELISM"] = "false"

import pandas as pd
import pdfplumber
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler("w9_extraction.log")],
)
logger = logging.getLogger(__name__)

CHECKBOX_LABELS = {
    0: "Individual/Sole Proprietor",
    1: "C Corporation",
    2: "S Corporation",
    3: "Partnership",
    4: "Trust/Estate",
    5: "LLC",
    6: "Other",
}

FIELD_MAP = {
    "f1_01": "name",
    "f1_02": "business_name",
    "f1_03": "llc_classification",
    "f1_04": "other_classification",
    "f1_05": "exempt_payee_code",
    "f1_06": "fatca_code",
    "f1_07": "address",
    "f1_08": "city_state_zip",
    "f1_09": "account_numbers",
    "f1_10": "requester_name_addr",
    "f1_11": "ssn_part1",
    "f1_12": "ssn_part2",
    "f1_13": "ssn_part3",
    "f1_14": "ein_part1",
    "f1_15": "ein_part2",
}

OUTPUT_COLUMNS = [
    "source_file", "name", "business_name", "tax_classification",
    "llc_classification", "address", "city_state_zip", "ssn", "ein",
    "exempt_payee_code", "fatca_code", "account_numbers", "date",
    "extraction_method", "extraction_notes",
]


def clean_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        try:
            value = value.decode("utf-8")
        except UnicodeDecodeError:
            value = value.decode("latin-1", errors="replace")
    return unicodedata.normalize("NFKC", str(value).strip())


def is_pdf_scanned(pdf_path: str) -> bool:
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = pdf.pages[0].extract_text() or ""
            return len(text.replace(" ", "").replace("\n", "")) < 50
    except Exception:
        return True


def has_acroform_fields(pdf_path: str) -> bool:
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                for annot in page.annots:
                    data = annot.get("data", {})
                    if "Tx" in str(data.get("FT", "")) and data.get("V"):
                        return True
        return False
    except Exception:
        return False


def extract_acroform(pdf_path: str) -> dict:
    record = {col: "" for col in OUTPUT_COLUMNS}
    raw_fields = {}
    checkbox_states = {}

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for annot in page.annots:
                title = annot.get("title", "") or ""
                data = annot.get("data", {})
                ft = str(data.get("FT", ""))
                v = data.get("V")
                ap_state = data.get("AS")

                if "Tx" in ft and v:
                    base = re.sub(r"\[\d+\]$", "", title)
                    raw_fields[base] = clean_text(v)
                elif "Btn" in ft and ap_state:
                    m = re.search(r"\[(\d+)\]", title)
                    if m:
                        idx = int(m.group(1))
                        checkbox_states[idx] = str(ap_state) not in ("/'Off'", "/Off", "Off")

    for field_key, col_name in FIELD_MAP.items():
        if field_key in raw_fields:
            record[col_name] = raw_fields[field_key]

    checked = [i for i, v in checkbox_states.items() if v]
    if checked:
        label = CHECKBOX_LABELS.get(checked[0], f"Unknown ({checked[0]})")
        if label == "LLC" and record.get("llc_classification"):
            llc_type = record["llc_classification"].upper()
            label = {"C": "LLC (C Corporation)", "S": "LLC (S Corporation)", "P": "LLC (Partnership)"}.get(llc_type, f"LLC ({llc_type})")
        record["tax_classification"] = label

    ssn = [raw_fields.get(k, "") for k in ["f1_11", "f1_12", "f1_13"]]
    if any(ssn):
        record["ssn"] = "-".join(p for p in ssn if p)

    ein = [raw_fields.get(k, "") for k in ["f1_14", "f1_15"]]
    if any(ein):
        record["ein"] = "-".join(p for p in ein if p)

    record["extraction_method"] = "AcroForm"
    return record


def _parse_w9_text(text: str, pdf_path: str) -> dict:
    record = {col: "" for col in OUTPUT_COLUMNS}
    record["source_file"] = Path(pdf_path).name
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    # Name: line immediately after the "Name of entity" header line
    record["name"] = _extract_name(lines)
    record["business_name"] = _extract_business_name(lines, text)
    record["tax_classification"] = _detect_classification(text)
    record["address"] = _extract_after_label(lines, r"^5\s+Address|Address\s*\(number")
    record["city_state_zip"] = _extract_after_label(lines, r"^6\s+City|City,\s*state")
    record["ssn"] = _extract_ssn(text)
    record["ein"] = _extract_ein(text)
    record["date"] = _extract_date(text, lines)
    return record


def _extract_name(lines: list) -> str:
    """Extract entity name from Line 1 of the W-9 form."""
    SKIP = re.compile(
        r"entity.s name on line|An entry is required|For a sole proprietor"
        r"|disregarded entity|enter the owner|enter the business"
        r"|Name\s+of\s+entity"
        r"|^(Check|Enter|See|Note:|Before|Give form|Business name)",
        re.IGNORECASE
    )
    found_label = False
    for line in lines:
        if re.search(r"Name\s+of\s+entity", line, re.IGNORECASE):
            found_label = True
            # Don't continue yet — the name might be on the SAME line after label text
            # (rare but possible in digital PDFs)
        if found_label:
            if SKIP.search(line):
                continue
            # Stop at next numbered field (line 2, 3a, etc.)
            if re.match(r"^[2-9]\s+[A-Z]", line):
                break
            if len(line) > 2 and re.match(r"[A-Za-z]", line):
                return line
    return ""


def _extract_after_label(lines: list, pattern: str) -> str:
    skip = [
        r"^(An entry|For a sole|Check the|Enter your|See instructions|Note:|If different|Business name|disregarded)",
        r"^\d+\s+[A-Z]",
    ]
    for i, line in enumerate(lines):
        if re.search(pattern, line, re.IGNORECASE):
            for j in range(i + 1, min(i + 6, len(lines))):
                c = lines[j]
                if not any(re.match(p, c, re.IGNORECASE) for p in skip) and len(c) > 2:
                    if not re.search(pattern, c, re.IGNORECASE):
                        return c
    return ""


def _extract_business_name(lines: list, text: str) -> str:
    for i, line in enumerate(lines):
        if re.search(r"^2\s+Business\s+name|Business\s+name.*?entity\s+name", line, re.IGNORECASE):
            for j in range(i + 1, min(i + 4, len(lines))):
                c = lines[j]
                if re.search(r"(different from above|disregarded entity|if different)", c, re.IGNORECASE):
                    continue
                if re.search(r"^(Check|Enter|See|Note:|3[ab]?|\d+\s+[A-Z])", c, re.IGNORECASE):
                    break
                if len(c) > 2:
                    return c

    m = re.search(
        r"Business\s+name.*?different\s+from\s+above\.?\s*\n+(.*?)(?:\n|3[ab]?\s)",
        text, re.IGNORECASE | re.DOTALL
    )
    if m:
        c = m.group(1).strip()
        if 2 < len(c) < 100 and not re.search(r"(Check|Enter|See )", c):
            return c
    return ""


def _detect_classification(text: str) -> str:
    """
    Detect which checkbox is ticked on line 3a.

    OCR renders a checkmark (✓) as various characters depending on the scan quality:
      Real checkmark chars : ✓ ✗ ☑ ■ ● ✔
      Tesseract misreads   : X x v [x] [v]
      Scanned form artefact: OCR sees the filled box as "im" or "wm" directly
                             before the label text on the same line.

    The line-3a row in a scanned W-9 looks like:
      "Oo individual/sole proprietor  im C corporation  Oo S corporation ..."
    where "Oo" = empty box, "im" = checked box artefact.

    Strategy: extract just the line-3a checkbox row, then test each label in
    priority order for the presence of a check token immediately before it.
    """
    check_tok = r"(?:✓|✗|☑|■|●|✔|\bim\b|\bwm\b|\[x\]|\[v\]|(?<![a-z])x(?![a-z])|(?<![a-z])v(?![a-z]))"

    # Find the checkbox row (line with "individual" and "corporation" on same line)
    checkbox_line = ""
    for line in text.splitlines():
        if re.search(r"individual.*corporation|corporation.*individual", line, re.IGNORECASE):
            checkbox_line = line
            break

    # Also grab the Other/LLC lines for those classifications
    other_lines = ""
    for line in text.splitlines():
        if re.search(r"Other\s*\(see|^.*\bLLC\b.*tax class", line, re.IGNORECASE):
            other_lines += " " + line

    classifications = [
        ("C Corporation",              r"C\s+corporation",           checkbox_line),
        ("S Corporation",              r"S\s+corporation",           checkbox_line),
        ("Partnership",                r"\bPartnership\b",          checkbox_line),
        ("Trust/Estate",               r"Trust.*?estate",             checkbox_line),
        ("Individual/Sole Proprietor", r"Individual.*?proprietor",    checkbox_line),
        ("LLC",                        r"\bLLC\b",                  other_lines + checkbox_line),
        ("Other",                      r"Other\s*\(see",            other_lines),
    ]
    for label, pattern, search_text in classifications:
        if not search_text:
            continue
        combined = rf"{check_tok}\s{{0,12}}{pattern}|{pattern}\s{{0,12}}{check_tok}"
        if re.search(combined, search_text, re.IGNORECASE):
            return label
    return ""


def _extract_ssn(text: str) -> str:
    # Standard format: 111-22-3333
    m = re.search(r"\b(\d{3})\s*[-]\s*(\d{2})\s*[-]\s*(\d{4})\b", text)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"

    # OCR spaced digits: "1 1 1 - 4 5 - 9 6 6 6" or single-box OCR
    ssn_area = re.search(r"Social\s*.?security\s*number(.*?)(?:or\s+Employer|Employer\s+ident)", text, re.IGNORECASE | re.DOTALL)
    if ssn_area:
        raw = ssn_area.group(1)
        digits = re.sub(r"[^\d]", "", raw)
        if len(digits) == 9:
            d = digits
            return f"{d[:3]}-{d[3:5]}-{d[5:]}"

    return ""


def _extract_ein(text: str) -> str:
    m = re.search(r"\b(\d{2})\s*[-\s]\s*(\d{7})\b", text)
    if m:
        return f"{m.group(1)}-{m.group(2)}"

    e = re.search(r"Employer\s+identification\s+number.*?(\d[\d\s\-]{6,12}\d)", text, re.IGNORECASE | re.DOTALL)
    if e:
        digits = re.sub(r"[^\d]", "", e.group(1))
        if len(digits) == 9:
            return f"{digits[:2]}-{digits[2:]}"
    return ""


def _extract_date(text: str, lines: list) -> str:
    # Look for date near signature section
    # Pattern: "Sign" ... date nearby
    for i, line in enumerate(lines):
        if re.search(r"^Sign\s|U\.S\.\s+person|Signature\s+of", line, re.IGNORECASE):
            # Check next few lines for a date
            for j in range(i, min(i + 5, len(lines))):
                m = re.search(r"(\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4})", lines[j])
                if m:
                    return m.group(1)

    # General date pattern in text
    m = re.search(r"\b(\d{1,2}[-/]\d{1,2}[-/]\d{4})\b", text)
    if m:
        return m.group(1)
    return ""


def extract_scanned_ocr(pdf_path: str) -> dict:
    record = {col: "" for col in OUTPUT_COLUMNS}
    try:
        from pdf2image import convert_from_path
        import pytesseract
        import cv2
        import numpy as np
    except ImportError as e:
        record["extraction_method"] = "OCR_FAILED"
        record["extraction_notes"] = f"Missing: {e}"
        return record

    try:
        images = convert_from_path(pdf_path, dpi=200)
        img = np.array(images[0])
        IMG_W, IMG_H = images[0].size
        SCALE = 200 / 72.0

        # OCR only top 75% of page — all W-9 data fields sit there,
        # cutting the signature/instruction block halves OCR time.
        h_px = img.shape[0]
        top_img = img[:int(h_px * 0.75), :]
        full_text = pytesseract.image_to_string(top_img, config="--oem 3 --psm 6")
        record = _parse_w9_text(full_text, pdf_path)

        # --- Date: targeted crop of the Sign Here / Date row ---
        # The Date row sits at ~70-78% page height; crop right column for the value.
        date_strip = img[int(h_px * 0.695): int(h_px * 0.755), int(IMG_W * 0.55):]
        if date_strip.size > 0:
            dg = cv2.cvtColor(date_strip, cv2.COLOR_RGB2GRAY)
            _, dp = cv2.threshold(dg, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            dp = cv2.resize(dp, (dp.shape[1]*3, dp.shape[0]*3), interpolation=cv2.INTER_CUBIC)
            date_raw = pytesseract.image_to_string(dp, config="--oem 3 --psm 7").strip()
            import re as _re
            dm = _re.search(r"(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})", date_raw)
            if dm:
                record["date"] = dm.group(1)

        # --- Targeted crop OCR for SSN and EIN ---
        # The digit cells are small and need isolated crop + whitelist for accuracy.
        # Anchor the search using the "Name of entity" label position.
        gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
        d = cv2.fastNlMeansDenoising(gray, h=10)
        thresh = cv2.adaptiveThreshold(d, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)

        # Import anchor finder from original extractor if available
        try:
            from w9_extractor import _find_form_anchor, _get_page_height_pt, _correct_ssn, _correct_ein
            pt_h = _get_page_height_pt(pdf_path)
            anchor = _find_form_anchor(thresh, SCALE, pt_h)
        except Exception:
            anchor = float(images[0].size[1]) / SCALE * 0.125  # fallback: 12.5% of page

        def crop_ocr_digits(x0, y_off, x1, h, wl):
            y0_pt = anchor + y_off
            px0 = max(0, int(x0 * SCALE)); py0 = max(0, int(y0_pt * SCALE))
            px1 = min(IMG_W, int(x1 * SCALE)); py1 = min(IMG_H, int((y0_pt + h) * SCALE))
            crop = img[py0:py1, px0:px1]
            if crop.size == 0:
                return ""
            g = cv2.cvtColor(crop, cv2.COLOR_RGB2GRAY)
            _, proc = cv2.threshold(g, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            proc = cv2.resize(proc, (proc.shape[1] * 4, proc.shape[0] * 4), interpolation=cv2.INTER_CUBIC)
            return pytesseract.image_to_string(
                proc, config=f'--oem 3 --psm 6 -c tessedit_char_whitelist="{wl}"'
            ).strip()

        # Targeted crop for SSN/EIN using absolute page y-coordinates.
        # Measured on real scans: SSN cells at abs ~385-435pt, EIN at ~420-455pt.
        # Scanning absolute coords (not anchor-relative) keeps the loop small (11 steps).
        def crop_ocr_abs(x0, y_abs, x1, h, wl):
            px0 = max(0, int(x0 * SCALE)); py0 = max(0, int(y_abs * SCALE))
            px1 = min(IMG_W, int(x1 * SCALE)); py1 = min(IMG_H, int((y_abs + h) * SCALE))
            crop = img[py0:py1, px0:px1]
            if crop.size == 0:
                return ""
            g = cv2.cvtColor(crop, cv2.COLOR_RGB2GRAY)
            _, proc = cv2.threshold(g, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            proc = cv2.resize(proc, (proc.shape[1] * 4, proc.shape[0] * 4), interpolation=cv2.INTER_CUBIC)
            return pytesseract.image_to_string(
                proc, config=f'--oem 3 --psm 6 -c tessedit_char_whitelist="{wl}"'
            ).strip()

        ssn_val = record.get("ssn", "")
        if not ssn_val:
            for y_abs in range(385, 440, 5):
                raw = crop_ocr_abs(410, y_abs, 575, 20, "0123456789- ")
                digits = re.sub(r"[^\d]", "", raw)
                if len(digits) == 9:
                    try:
                        ssn_val = _correct_ssn(raw)
                    except Exception:
                        ssn_val = f"{digits[:3]}-{digits[3:5]}-{digits[5:]}"
                    if ssn_val:
                        break
            record["ssn"] = ssn_val

        ein_val = record.get("ein", "")
        if not ein_val:
            for y_abs in range(420, 460, 5):
                raw = crop_ocr_abs(410, y_abs, 575, 20, "0123456789-")
                digits = re.sub(r"[^\d]", "", raw)
                if len(digits) == 9:
                    try:
                        ein_val = _correct_ein(raw)
                    except Exception:
                        ein_val = f"{digits[:2]}-{digits[2:]}"
                    if ein_val:
                        break
            record["ein"] = ein_val

        record["extraction_method"] = "OCR"
    except Exception as e:
        logger.error(f"OCR failed: {e}")
        record["extraction_notes"] = str(e)
        record["extraction_method"] = "OCR_FAILED"
    return record


def extract_text_based(pdf_path: str) -> dict:
    with pdfplumber.open(pdf_path) as pdf:
        text = "".join((p.extract_text(layout=True) or "") + "\n" for p in pdf.pages[:2])
    record = _parse_w9_text(text, pdf_path)
    record["extraction_method"] = "TextPDF"
    return record


def extract_w9(pdf_path: str) -> dict:
    record = {col: "" for col in OUTPUT_COLUMNS}
    record["source_file"] = Path(pdf_path).name
    logger.info(f"Processing: {Path(pdf_path).name}")
    try:
        if has_acroform_fields(pdf_path):
            data = extract_acroform(pdf_path)
        elif is_pdf_scanned(pdf_path):
            data = extract_scanned_ocr(pdf_path)
        else:
            data = extract_text_based(pdf_path)
        record.update(data)
        record["source_file"] = Path(pdf_path).name
    except Exception as e:
        logger.error(f"Failed: {e}")
        record["extraction_notes"] = str(e)
        record["extraction_method"] = "FAILED"
    return record


def save_to_excel(records: list, output_path: str) -> str:
    df = pd.DataFrame(records, columns=OUTPUT_COLUMNS)
    client_columns = {
        "name": "Name of Entity(1)",
        "business_name": "Business Name(2)",
        "tax_classification": "TOB(3a)",
        "address": "Address(5)",
        "city_state_zip": "City, state, and ZIP code(6)",
        "ssn": "Social Security Number",
        "ein": "Employer Identification Number",
        "date": "Date",
    }
    df_out = df[list(client_columns.keys())].rename(columns=client_columns)
    out = output_path.replace(".xlsm", ".xlsx")
    df_out.to_excel(out, index=False, sheet_name="W-9 Data")

    wb = load_workbook(out)
    ws = wb.active

    hf = PatternFill(fill_type="solid", start_color="1F4E79", end_color="1F4E79")
    for cell in ws[1]:
        cell.fill = hf
        cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    ws.row_dimensions[1].height = 35

    lf = PatternFill(fill_type="solid", start_color="EBF0FA", end_color="EBF0FA")
    for row in range(2, ws.max_row + 1):
        for cell in ws[row]:
            if row % 2 == 0:
                cell.fill = lf
            cell.font = Font(name="Arial", size=9)
            cell.alignment = Alignment(vertical="center")
            cell.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
        ws.row_dimensions[row].height = 18

    col_widths = {"Name of Entity(1)": 30, "Business Name(2)": 30, "TOB(3a)": 28,
                  "Address(5)": 30, "City, state, and ZIP code(6)": 25,
                  "Social Security Number": 22, "Employer Identification Number": 28, "Date": 18}
    for i, col in enumerate(client_columns.values(), 1):
        ws.column_dimensions[get_column_letter(i)].width = col_widths.get(col, 20)

    ws.freeze_panes = "A2"
    wb.save(out)
    logger.info(f"Saved {len(records)} records to {out}")
    return out


def run_pipeline(input_dir="input_pdfs", output_excel="w9-extractor.xlsx",
                 processed_dir="output_pdfs", move_files=True):
    input_path = Path(input_dir)
    if not input_path.exists():
        logger.error(f"Not found: {input_dir}")
        return

    pdf_files = sorted(input_path.glob("*.pdf"))
    if not pdf_files:
        logger.warning(f"No PDFs in: {input_dir}")
        return

    if move_files:
        Path(processed_dir).mkdir(parents=True, exist_ok=True)

    records = []
    for pdf_path in pdf_files:
        record = extract_w9(str(pdf_path))
        records.append(record)
        if move_files:
            name = record.get("name", "Unknown")
            safe = re.sub(r'[<>:"/\\|?*]', "_", name)[:80] or "Unknown"
            dest = Path(processed_dir) / f"{safe}_{pdf_path.stem}.pdf"
            try:
                shutil.copy2(pdf_path, dest)
                os.remove(pdf_path)
            except Exception as e:
                logger.warning(f"Move failed: {e}")

    if records:
        save_to_excel(records, output_excel)
        print(f"Done. {len(records)} records processed.")


if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument("--input", default="input_pdfs")
    p.add_argument("--output", default="w9-extractor.xlsx")
    p.add_argument("--archive", default="output_pdfs")
    p.add_argument("--no-move", action="store_true")
    args = p.parse_args()
    run_pipeline(args.input, args.output, args.archive, not args.no_move)
