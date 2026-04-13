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
    
    text = str(value)

    text = re.sub(r'[_{}\[\]|]', ' ', text) 
    return unicodedata.normalize("NFKC", text.strip())

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
    record["tax_classification"] = _detect_classification(text, record["name"])
    record["address"] = _extract_after_label(lines, r"^5\s+Address|Address\s*\(number")
    record["city_state_zip"] = _extract_after_label(lines, r"^6\s+City|City,\s*state")
    record["ssn"] = _extract_ssn(text)
    record["ein"] = _extract_ein(text)
    record["date"] = _extract_date(text, lines)
    return record


def _extract_name(lines: list) -> str:
    # Strict list of headers to ignore
    SKIP = re.compile(
        r"Form W-9|Request for Taxpayer|Department of the Treasury|Internal Revenue Service"
        r"|entity.?s name on|line 2|An entry is required|sole proprietor"
        r"|Name\s+of\s+entity|1\s+Name|Go to www\.irs\.gov",
        re.IGNORECASE
    )

    for i, line in enumerate(lines):
        line = line.strip()
        # Look specifically for the "1 Name" anchor
        if re.search(r"1\s*Name\s*of\s*entity", line, re.IGNORECASE):
            # Check the following lines for the actual value
            for j in range(i + 1, i + 3):
                if j < len(lines):
                    candidate = lines[j].strip()
                    if candidate and not SKIP.search(candidate) and len(candidate) > 2:
                        return candidate
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


# def clean_text(text: str) -> str:
#     if not text:
#         return ""
#     # Remove OCR noise like |, {, [, ], _, or dots at the start/end
#     cleaned = re.sub(r'^[^a-zA-Z0-9]+', '', text)
#     return cleaned.strip()

def _clean_name_noise(text: str) -> str:
    """Очищает строку от типичного мусора OCR, сохраняя валидный текст."""
    if not text:
        return ""
    # Заменяем _, {, }, [, ], |, и обратные слеши на пробелы
    cleaned = re.sub(r'[_{}\[\]|\\]', ' ', text)
    # Схлопываем множественные пробелы в один
    cleaned = re.sub(r'\s+', ' ', cleaned)
    # Удаляем не-буквенно-цифровые символы в самом начале и конце строки
    cleaned = re.sub(r'^[^a-zA-Z0-9]+|[^a-zA-Z0-9]+$', '', cleaned)
    return cleaned.strip()

def _extract_business_name(lines: list, text: str) -> str:
    for i, line in enumerate(lines):
        # Anchor for Line 2
        if re.search(r"2\s*Business\s*name", line, re.IGNORECASE):
            # Check the same line first (after the label)
            if "different from above" in line.lower():
                parts = re.split(r"above\.?", line, flags=re.IGNORECASE)
                if len(parts) > 1:
                    candidate = clean_text(parts[1])
                    candidate = _clean_name_noise(candidate) # Применяем очистку
                    if len(candidate) > 2:
                        return candidate
            
            # Check the next 2 lines
            for j in range(i + 1, i + 3):
                if j < len(lines):
                    candidate = clean_text(lines[j])
                    candidate = _clean_name_noise(candidate) # Применяем очистку
                    # Ensure it's not the start of Line 3
                    if candidate and not re.search(r"^(3a|Check|3b)", candidate, re.IGNORECASE):
                        if len(candidate) > 2:
                            return candidate
    return ""

def _detect_classification(text: str, name: str = "") -> str:
    text = (text or "").lower()
    name = (name or "").lower()

    # Explicit classification labels found on the form are highest priority.
    if re.search(r'\bc\s*(corporat\w*|corp|corp\.)\b', text):
        return "C Corporation"
    if re.search(r'\bs\s*(corporat\w*|corp|corp\.)\b', text):
        return "S Corporation"
    if re.search(r'\bpartnership\b', text):
        return "Partnership"
    if re.search(r'\btrust/?estate\b', text):
        return "Trust/Estate"
    if re.search(r'\bother\b', text):
        return "Other"
    if re.search(r'\bindividual/sole proprietor\b|\bsole proprietor\b|\bindividual\b', text):
        return "Individual/Sole Proprietor"
    if re.search(r'\bllc\b', text):
        # If the form also contains a more specific corporate selection, honor that instead.
        if re.search(r'\bc\s*(corporation|orp|orp\.)\b', text):
            return "C Corporation"
        if re.search(r'\bs\s*(corporation|orp|orp\.)\b', text):
            return "S Corporation"
        if re.search(r'\bpartnership\b', text):
            return "Partnership"
        return "LLC"

    if any(x in name for x in ["inc", "corp", "corporation"]):
        return "C Corporation"
    if any(x in name for x in ["llp", "partnership"]):
        return "Partnership"
    if "llc" in name:
        return "LLC"

    return "Individual/Sole Proprietor"


def _extract_ssn_from_image(img) -> str:
    try:
        import cv2
        import pytesseract
        import numpy as np
    except ImportError:
        return ""

    h, w = img.shape[:2]
    candidate_regions = [
        (int(w * 0.60), int(h * 0.42), min(w, int(w * 0.98)), int(h * 0.52)),
        (int(w * 0.62), int(h * 0.40), min(w, int(w * 0.98)), int(h * 0.55)),
        (int(w * 0.58), int(h * 0.38), min(w, int(w * 0.98)), int(h * 0.56)),
    ]

    for x0, y0, x1, y1 in candidate_regions:
        crop = img[y0:y1, x0:x1]
        if crop.size == 0:
            continue
        gray = cv2.cvtColor(crop, cv2.COLOR_RGB2GRAY)
        _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        proc = cv2.bitwise_not(thresh)
        proc = cv2.resize(proc, (proc.shape[1] * 3, proc.shape[0] * 3), interpolation=cv2.INTER_CUBIC)
        text = pytesseract.image_to_string(proc, config='--oem 3 --psm 6 -c tessedit_char_whitelist="0123456789- "')
        match = re.search(r'(\d{3})\D*(\d{2})\D*(\d{4})', text)
        if match:
            return f"{match.group(1)}-{match.group(2)}-{match.group(3)}"
    return ""


def _extract_ein_from_image(img) -> str:
    try:
        import cv2
        import pytesseract
        import numpy as np
    except ImportError:
        return ""

    h, w = img.shape[:2]
    candidate_regions = [
        (int(w * 0.60), int(h * 0.50), min(w, int(w * 0.98)), int(h * 0.60)),
        (int(w * 0.58), int(h * 0.48), min(w, int(w * 0.98)), int(h * 0.62)),
    ]

    for x0, y0, x1, y1 in candidate_regions:
        crop = img[y0:y1, x0:x1]
        if crop.size == 0:
            continue
        gray = cv2.cvtColor(crop, cv2.COLOR_RGB2GRAY)
        _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        proc = cv2.bitwise_not(thresh)
        proc = cv2.resize(proc, (proc.shape[1] * 3, proc.shape[0] * 3), interpolation=cv2.INTER_CUBIC)
        text = pytesseract.image_to_string(proc, config='--oem 3 --psm 6 -c tessedit_char_whitelist="0123456789- "')
        match = re.search(r'(\d{2})\D*(\d{7})', text)
        if match:
            return f"{match.group(1)}-{match.group(2)}"
    return ""


def _extract_ssn(text: str) -> str:
    """Extract SSN in various formats (XXX-XX-XXXX, XXX XX XXXX, XXXXXXXXX, etc)."""
    if not text:
        return ""
    
    # Clean OCR noise but keep structure
    clean_text = re.sub(r'[_{}\[\]|\\]', ' ', text)
    
    # Pattern 1: Standard format (XXX-XX-XXXX or XXX XX XXXX or with dots)
    m = re.search(r'\b(\d{3})\s*[-.\s]\s*(\d{2})\s*[-.\s]\s*(\d{4})\b', clean_text)
    if m:
        result = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
        logger.info(f"SSN found via Pattern 1: {result}")
        return result
    
    # Pattern 2: Look for various SSN labels and capture digits after
    m_label = re.search(
        r"social\s+security\s+number[:\s]*([\d\s\-\.\|]{9,40})",
        clean_text,
        re.IGNORECASE,
    )
    if m_label:
        digits = re.sub(r"[^\d]", "", m_label.group(1))
        if len(digits) >= 9:
            result = f"{digits[:3]}-{digits[3:5]}-{digits[5:9]}"
            logger.info(f"SSN found via Pattern 2 (label): {result}")
            return result

    # Pattern 2b: Look for the shorter label 'SSN' as a fallback
    m_ssn_label = re.search(r"\bSSN\b[:\s]*([\d\s\-\.]{9,30})", clean_text, re.IGNORECASE)
    if m_ssn_label:
        digits = re.sub(r"[^\d]", "", m_ssn_label.group(1))
        if len(digits) >= 9:
            result = f"{digits[:3]}-{digits[3:5]}-{digits[5:9]}"
            logger.info(f"SSN found via Pattern 2b (SSN): {result}")
            return result
    
    # Pattern 3: Look for 9 consecutive digits (with minimal separators)
    for m in re.finditer(r'(\d{3})\s*[-.\s]*(\d{2})\s*[-.\s]*(\d{4})', clean_text):
        area, group, serial = m.group(1), m.group(2), m.group(3)
        # Skip invalid SSNs: area 000, 666, or 900-999
        if area not in ['000', '666'] and not (900 <= int(area) <= 999):
            result = f"{area}-{group}-{serial}"
            logger.info(f"SSN found via Pattern 3: {result}")
            return result
    
    # Pattern 4: Fall back to any 9-digit sequence
    m_fallback = re.search(r'\b(\d{9})\b', clean_text)
    if m_fallback:
        digits = m_fallback.group(1)
        result = f"{digits[:3]}-{digits[3:5]}-{digits[5:9]}"
        logger.info(f"SSN found via Pattern 4 (fallback): {result}")
        return result
    
    # Log OCR extraction findings for debugging
    logger.info(f"No SSN found. Text sample: {clean_text[:200]}")
    
    return ""

def _extract_ein(text: str) -> str:
    """Extract EIN in various formats (XX-XXXXXXX, XX XXXXXXX, XXXXXXXXX, etc)."""
    if not text:
        return ""
    
    # Clean OCR noise but keep structure
    clean_text = re.sub(r'[_{}\[\]|\\]', ' ', text)
    
    # Pattern 1: Standard format (XX-XXXXXXX or XX XXXXXXX or with dots)
    m = re.search(r'\b(\d{2})\s*[-.\s]\s*(\d{7})\b', clean_text)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    
    # Pattern 2: Look for "Employer Identification Number" label and capture digits after
    m_label = re.search(r"Employer\s+Identification\s+Number[:\s]+([\d\s\-\.]{7,20})", clean_text, re.IGNORECASE)
    if m_label:
        digits = re.sub(r"[^\d]", "", m_label.group(1))
        if len(digits) >= 9:
            return f"{digits[:2]}-{digits[2:9]}"
    
    # Pattern 3: Look for 9 consecutive digits
    for m in re.finditer(r'(\d{2})\s*[-.\s]*(\d{7})', clean_text):
        return f"{m.group(1)}-{m.group(2)}"
    
    # Pattern 4: Fall back to any 9-digit sequence for EIN
    m_fallback = re.search(r'\b(\d{9})\b', clean_text)
    if m_fallback:
        digits = m_fallback.group(1)
        return f"{digits[:2]}-{digits[2:9]}"
    
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
        images = convert_from_path(pdf_path, dpi=200, first_page=1, last_page=1)
        img = np.array(images[0])
        IMG_W, IMG_H = images[0].size
        SCALE = 200 / 72.0

        logger.info("OCR processing only first page")
        h_px = img.shape[0]
        top_img = img[:int(h_px * 0.75), :]
        full_text = pytesseract.image_to_string(top_img, config="--oem 3 --psm 6")
        record = _parse_w9_text(full_text, pdf_path)

        # fill missing OCR-only fields from the image region if the parsed text did not include them
        if not record.get("ssn"):
            ssn_found = _extract_ssn_from_image(img)
            if ssn_found:
                record["ssn"] = ssn_found
        if not record.get("ein"):
            ein_found = _extract_ein_from_image(img)
            if ein_found:
                record["ein"] = ein_found

        # --- Date: crop the signature section at bottom right ---
        # Look in the bottom 20% of the page, right half for date patterns
        date_strip = img[int(h_px * 0.8):, int(IMG_W * 0.5):]
        if date_strip.size > 0:
            dg = cv2.cvtColor(date_strip, cv2.COLOR_RGB2GRAY)
            _, dp = cv2.threshold(dg, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            dp = cv2.resize(dp, (dp.shape[1]*4, dp.shape[0]*4), interpolation=cv2.INTER_CUBIC)
            date_raw = pytesseract.image_to_string(dp, config="--oem 3 --psm 6").strip()
            import re as _re
            dm = _re.search(r"(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})", date_raw)
            if dm:
                record["date"] = dm.group(1)

        record['extraction_method'] = 'OCR'
    except Exception as e:
        logger.error(f"OCR failed: {e}")
        record["extraction_notes"] = str(e)
        record["extraction_method"] = "OCR_FAILED"
    return record


def extract_text_based(pdf_path: str) -> dict:
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        text = page.extract_text(layout=True) or ""
    record = _parse_w9_text(text, pdf_path)
    record["extraction_method"] = "TextPDF"
    return record


def extract_w9(pdf_path: str) -> dict:
    record = {col: "" for col in OUTPUT_COLUMNS}
    record["source_file"] = Path(pdf_path).name

    logger.info(f"Processing: {record['source_file']}")

    try:
        if has_acroform_fields(pdf_path):
            method = "AcroForm"
            data = extract_acroform(pdf_path)

        elif is_pdf_scanned(pdf_path):
            method = "OCR"
            data = extract_scanned_ocr(pdf_path)

        else:
            method = "TextPDF"
            data = extract_text_based(pdf_path)

        for k, v in data.items():
            if v: 
                record[k] = v

        record["extraction_method"] = data.get("extraction_method", method)
        record["extraction_status"] = "SUCCESS"

    except Exception as e:
        logger.error(f"Failed: {e}")
        record["extraction_notes"] = str(e)
        record["extraction_method"] = "FAILED"
        record["extraction_status"] = "FAILED"

    if not any([
        record.get("name"),
        record.get("ssn"),
        record.get("ein")
    ]):
        record["extraction_notes"] = (record.get("extraction_notes", "") + " | EMPTY_RECORD").strip()
        if record["extraction_status"] != "FAILED":
            record["extraction_status"] = "EMPTY"

    logger.info(
        f"{record['source_file']} | "
        f"method={record.get('extraction_method')} | "
        f"status={record.get('extraction_status')} | "
        f"name={record.get('name')} | "
        f"ssn={record.get('ssn')} | "
        f"ein={record.get('ein')}"
    )

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
