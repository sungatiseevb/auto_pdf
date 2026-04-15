"""
W-9 extraction utility for IRS Form W-9.
This version uses PyMuPDF widgets first, then full text extraction, and finally OCR for scanned forms.
It returns a clean dictionary ready for Excel export.
"""

import logging
import re
import traceback
import unicodedata
from pathlib import Path
from typing import Dict, Optional

try:
    import fitz  # PyMuPDF
except ImportError:  # pragma: no cover
    fitz = None

try:
    from pdf2image import convert_from_path
except ImportError:  # pragma: no cover
    convert_from_path = None

try:
    import pytesseract
except ImportError:  # pragma: no cover
    pytesseract = None

try:
    import cv2
except ImportError:  # pragma: no cover
    cv2 = None

try:
    import numpy as np
except ImportError:  # pragma: no cover
    np = None


LOG_PATH = "w9_extraction.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler(LOG_PATH), logging.StreamHandler()],
)
logger = logging.getLogger(__name__)


OUTPUT_FIELDS = [
    "filename",
    "name_entity",
    "business_name",
    "tob",
    "address",
    "city_state_zip",
    "ssn",
    "ein",
    "date",
    "status",
    "method",
    "raw_snippet",
]

WIDGET_FIELD_MAP = {
    "f1_01": "name_entity",
    "f1_02": "business_name",
    "f1_07": "address",
    "f1_08": "city_state_zip",
    "f1_11": "ssn_part1",
    "f1_12": "ssn_part2",
    "f1_13": "ssn_part3",
    "f1_14": "ein_part1",
    "f1_15": "ein_part2",
}

CHECKBOX_PATTERNS = {
    "Individual/sole proprietor": [r"individual/?sole proprietor", r"individual sole proprietor"],
    "C corporation": [r"c(?:\s|-)corporation", r"c corporation", r"c corp"],
    "S corporation": [r"s(?:\s|-)corporation", r"s corporation", r"s corp"],
    "Partnership": [r"partnership"],
    "Trust/estate": [r"trust/?estate", r"trust estate"],
    "Other": [r"\bother\b"],
}

DATE_PATTERN = re.compile(r"(?P<m>\d{1,2})[\./-](?P<d>\d{1,2})[\./-](?P<y>\d{2,4})")
SSN_PATTERN = re.compile(r"\b(\d{3})[-\s]*(\d{2})[-\s]*(\d{4})\b")
EIN_PATTERN = re.compile(r"\b(\d{2})[-\s]*(\d{7})\b")

OCR_WHITELIST = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz.,-/() "


def _normalize_text(value: Optional[str]) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        value = value.decode("utf-8", errors="ignore")
    text = str(value)
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\r", "\n")
    text = re.sub(r"[\t\x0b\x0c]+", " ", text)
    return text.strip()


def _clean_line(value: str) -> str:
    text = _normalize_text(value)
    text = re.sub(r"[_{}\[\]|\\]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _normalize_date(value: str) -> Optional[str]:
    text = _normalize_text(value)
    match = DATE_PATTERN.search(text)
    if not match:
        return None
    month = int(match.group("m"))
    day = int(match.group("d"))
    year = match.group("y")
    if len(year) == 2:
        year = int(year)
        year = 2000 + year if year < 80 else 1900 + year
    else:
        year = int(year)
    if not (1 <= month <= 12 and 1 <= day <= 31):
        return None
    return f"{month:02d}/{day:02d}/{year:04d}"


def _extract_ssn(text: str) -> Optional[str]:
    if not text:
        return None
    candidate = _normalize_text(text)
    match = SSN_PATTERN.search(candidate)
    if match:
        return f"{match.group(1)}-{match.group(2)}-{match.group(3)}"
    fallback = re.search(r"\b(\d{3})\s+(\d{2})\s+(\d{4})\b", candidate)
    if fallback:
        return f"{fallback.group(1)}-{fallback.group(2)}-{fallback.group(3)}"
    return None


def _extract_ein(text: str) -> Optional[str]:
    if not text:
        return None
    candidate = _normalize_text(text)
    match = EIN_PATTERN.search(candidate)
    if match:
        return f"{match.group(1)}-{match.group(2)}"
    fallback = re.search(r"\b(\d{2})\s+(\d{7})\b", candidate)
    if fallback:
        return f"{fallback.group(1)}-{fallback.group(2)}"
    return None


def _is_label_remainder(value: str) -> bool:
    if not value:
        return False
    remainder = value.lower().strip()
    remainder = re.sub(r"^[\W_]+", "", remainder)
    return bool(
        re.search(
            r"^(an entry is required|for a sole proprietor|if different from above|line \d+|entitys? name|business name|address|number, street|apt\.|suite no|city, state|city.*state|state.*zip|and zip code|zip code|see instructions|enter the|requester|purpose of form|disregarded entity|if different|office number|name of entity|requesters name)",
            remainder,
        )
    )


def _find_label_value(lines: list, label_patterns: list, skip_patterns: Optional[list] = None) -> Optional[str]:
    skip_regex = [re.compile(pattern, re.IGNORECASE) for pattern in (skip_patterns or [])]
    for index, line in enumerate(lines):
        normalized = line.strip()
        for pattern in label_patterns:
            match = re.search(pattern, normalized, re.IGNORECASE)
            if not match:
                continue
            after = normalized[match.end():].strip(" :.-")
            if after:
                after = re.sub(r"^[\./\-\s]+", "", after)
            if after and len(after) > 1 and not _is_label_remainder(after):
                return after
            for next_line in lines[index + 1 : index + 7]:
                if not next_line:
                    continue
                if any(regex.search(next_line) for regex in skip_regex):
                    continue
                candidate = next_line.strip()
                candidate = re.sub(r"^[\./\-\s]+", "", candidate)
                if candidate and len(candidate) > 1 and not _is_label_remainder(candidate):
                    return candidate
    return None


def _extract_name_entity(lines: list, text: str) -> Optional[str]:
    result = _find_label_value(
        lines,
        [r"1\s*Name of entity/individual", r"1\s*Name of entity", r"Name of entity/individual", r"Name of entity"],
        skip_patterns=[r"^An entry is required", r"^For a sole proprietor", r"^\(For a sole proprietor"],
    )
    if result:
        return result
    for index, line in enumerate(lines[:6]):
        if re.search(r"^\d+\s*Name", line, re.IGNORECASE):
            next_line = lines[index + 1] if index + 1 < len(lines) else ""
            if next_line and not re.search(r"^\d+", next_line):
                return next_line.strip()
    return None


def _extract_business_name(lines: list) -> Optional[str]:
    return _find_label_value(
        lines,
        [r"2\s*Business name", r"Business name", r"disregarded entity name"],
        skip_patterns=[r"^if different from above", r"^An entry is required"],
    )


def _extract_address(lines: list) -> Optional[str]:
    return _find_label_value(
        lines,
        [r"\b5\s*Address", r"\b5\s*Address\b", r"Address\s*\(number", r"^Address"],
        skip_patterns=[r"^See instructions", r"^City"],
    )


def _extract_city_state_zip(lines: list) -> Optional[str]:
    return _find_label_value(
        lines,
        [r"\b6\s*City", r"City,\s*state", r"City.*state.*ZIP", r"City and state"],
    )


def _extract_tob(text: str, lines: list) -> Optional[str]:
    candidate = _normalize_text(text).lower()
    if re.search(r"llc\s*[-=]\s*c|c\s*=\s*llc|llc.*c\s*=", candidate):
        return "LLC - C"
    if re.search(r"llc\s*[-=]\s*s|s\s*=\s*llc|llc.*s\s*=", candidate):
        return "LLC - S"
    if re.search(r"llc\s*[-=]\s*p|p\s*=\s*llc|llc.*p\s*=", candidate):
        return "LLC - P"
    for label, patterns in CHECKBOX_PATTERNS.items():
        for pattern in patterns:
            if re.search(pattern, candidate):
                return label
    if re.search(r"\bllc\b", candidate):
        return "LLC"
    return None


def _extract_date(text: str, lines: list) -> Optional[str]:
    for index, line in enumerate(lines):
        if re.search(r"signature|sign\b|date\b|part\s*ii", line, re.IGNORECASE):
            for search_line in lines[index : index + 6]:
                normalized = _normalize_text(search_line)
                date_value = _normalize_date(normalized)
                if date_value:
                    return date_value
    for match in DATE_PATTERN.finditer(text):
        date_value = _normalize_date(match.group(0))
        if date_value:
            return date_value
    return None


def _parse_text_content(text: str) -> Dict[str, Optional[str]]:
    if text is None:
        text = ""
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    return {
        "filename": None,
        "name_entity": _extract_name_entity(lines, text),
        "business_name": _extract_business_name(lines),
        "tob": _extract_tob(text, lines),
        "address": _extract_address(lines),
        "city_state_zip": _extract_city_state_zip(lines),
        "ssn": _extract_ssn(text),
        "ein": _extract_ein(text),
        "date": _extract_date(text, lines),
        "status": None,
        "method": None,
        "raw_snippet": text[:800] if text else None,
    }


def _widgets_available() -> bool:
    return fitz is not None


def _extract_from_widgets(pdf_path: str) -> Dict[str, Optional[str]]:
    result = {
        "name_entity": None,
        "business_name": None,
        "tob": None,
        "address": None,
        "city_state_zip": None,
        "ssn": None,
        "ein": None,
        "date": None,
        "raw_snippet": None,
    }
    try:
        if not _widgets_available():
            return result
        with fitz.open(pdf_path) as doc:
            values = {}
            for page in doc:
                widgets = page.widgets() or []
                for widget in widgets:
                    name = getattr(widget, "field_name", None) or getattr(widget, "name", None)
                    value = getattr(widget, "field_value", None) or getattr(widget, "text", None)
                    if not name or value is None:
                        continue
                    key = _normalize_text(name).lower()
                    value_text = _clean_line(value)
                    if not value_text:
                        continue
                    if key in WIDGET_FIELD_MAP:
                        values[WIDGET_FIELD_MAP[key]] = value_text
                        continue
                    if "ssn" in key:
                        values["ssn"] = _extract_ssn(value_text)
                        continue
                    if "ein" in key:
                        values["ein"] = _extract_ein(value_text)
                        continue
                    if "name" in key and "business" not in key and not result["name_entity"]:
                        result["name_entity"] = value_text
                        continue
                    if "business" in key and not result["business_name"]:
                        result["business_name"] = value_text
                        continue
                    if "address" in key and not result["address"]:
                        result["address"] = value_text
                        continue
                    if "city" in key and not result["city_state_zip"]:
                        result["city_state_zip"] = value_text
                        continue
                    if "date" in key and not result["date"]:
                        result["date"] = _normalize_date(value_text)
                        continue
            if values.get("ssn_part1") and values.get("ssn_part2") and values.get("ssn_part3"):
                result["ssn"] = f"{values['ssn_part1']}-{values['ssn_part2']}-{values['ssn_part3']}"
            if values.get("ein_part1") and values.get("ein_part2"):
                result["ein"] = f"{values['ein_part1']}-{values['ein_part2']}"
            for field_name, field_value in values.items():
                if field_name.startswith("ssn_part") or field_name.startswith("ein_part"):
                    continue
                if field_value:
                    result[field_name] = field_value
            result["raw_snippet"] = None
            return result
    except Exception:
        logger.debug("Widget extraction failed:\n%s", traceback.format_exc())
        return result


def _is_text_scanned(text: str) -> bool:
    if not text:
        return True
    words = re.findall(r"\w+", text)
    return len(words) < 40


def _extract_text_from_pdf(pdf_path: str) -> str:
    try:
        if fitz is None:
            return ""
        with fitz.open(pdf_path) as doc:
            page_text = []
            for page in doc:
                page_text.append(page.get_text("text") or "")
            return "\n".join(page_text).strip()
    except Exception:
        logger.debug("Text extraction failed:\n%s", traceback.format_exc())
        return ""


def _preprocess_ocr_image(image: "np.ndarray") -> Optional["np.ndarray"]:
    if cv2 is None or np is None:
        return None
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
    blurred = cv2.medianBlur(gray, 3)
    _, thresh = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return thresh


def _extract_ocr_text(pdf_path: str) -> str:
    if convert_from_path is None or pytesseract is None or cv2 is None or np is None:
        logger.warning("OCR dependencies are missing.")
        return ""
    try:
        images = convert_from_path(pdf_path, dpi=350, first_page=1, last_page=1)
        if not images:
            return ""
        image = np.array(images[0])
        processed = _preprocess_ocr_image(image)
        if processed is None:
            return ""
        config = f'--oem 3 --psm 6 -c tessedit_char_whitelist="{OCR_WHITELIST}"'
        text = pytesseract.image_to_string(processed, config=config)
        if not text.strip():
            inverted = cv2.bitwise_not(processed)
            text = pytesseract.image_to_string(inverted, config=config)
        return text.strip()
    except Exception:
        logger.debug("OCR extraction failed:\n%s", traceback.format_exc())
        return ""


def _get_ocr_words(image: "np.ndarray") -> list[dict]:
    if pytesseract is None:
        return []
    try:
        data = pytesseract.image_to_data(image, config='--oem 3 --psm 6', output_type=pytesseract.Output.DICT)
    except Exception:
        return []
    words = []
    for i, text in enumerate(data.get("text", [])):
        if not text or not text.strip():
            continue
        words.append(
            {
                "text": text.strip(),
                "left": data.get("left", [])[i],
                "top": data.get("top", [])[i],
                "width": data.get("width", [])[i],
                "height": data.get("height", [])[i],
            }
        )
    return words


def _group_words_by_line(words: list[dict], y_tolerance: int = 10) -> list[tuple[int, list[dict]]]:
    lines = []
    words_sorted = sorted(words, key=lambda w: (w["top"], w["left"]))
    current_line = []
    current_top = None
    for word in words_sorted:
        if current_top is None:
            current_top = word["top"]
        if abs(word["top"] - current_top) > y_tolerance:
            lines.append((current_top, current_line))
            current_line = []
            current_top = word["top"]
        current_line.append(word)
    if current_line:
        lines.append((current_top, current_line))
    return lines


def _extract_number_near_label(image: "np.ndarray", label_patterns: list[str], required_digits: int = 9) -> Optional[str]:
    words = _get_ocr_words(image)
    if not words:
        return None
    lines = _group_words_by_line(words)
    for index, (_, line_words) in enumerate(lines):
        line_text = " ".join(word["text"] for word in line_words)
        for label in label_patterns:
            if re.search(label, line_text, re.IGNORECASE):
                combined = line_text
                if index + 1 < len(lines):
                    combined += " " + " ".join(w["text"] for w in lines[index + 1][1])
                digits = re.sub(r"[^0-9]", "", combined)
                if len(digits) >= required_digits:
                    return digits[:required_digits]
                break
    return None


def _merge_records(primary: Dict[str, Optional[str]], secondary: Dict[str, Optional[str]]) -> Dict[str, Optional[str]]:
    merged = primary.copy()
    for key, value in secondary.items():
        if key in ("filename", "status", "method"):
            continue
        if not merged.get(key) and value:
            merged[key] = value
    if not merged.get("raw_snippet"):
        merged["raw_snippet"] = secondary.get("raw_snippet")
    return merged


def _determine_status(record: Dict[str, Optional[str]]) -> str:
    required_fields = ["name_entity", "tob", "address", "city_state_zip"]
    if record.get("status") == "ERROR":
        return "ERROR"
    if all(record.get(field) for field in required_fields) and (record.get("ssn") or record.get("ein")):
        return "SUCCESS"
    if any(record.get(field) for field in OUTPUT_FIELDS if field not in {"status", "method", "raw_snippet", "filename"}):
        return "PARTIAL"
    return "NOT_FOUND"


def extract_w9(pdf_path: str) -> Dict[str, Optional[str]]:
    result = {field: None for field in OUTPUT_FIELDS}
    result["filename"] = Path(pdf_path).name
    result["method"] = "text"
    result["status"] = "NOT_FOUND"
    text_source = ""
    try:
        widget_data = _extract_from_widgets(pdf_path)
        has_widget_data = any(widget_data.get(field) for field in ["name_entity", "business_name", "address", "city_state_zip", "ssn", "ein"])
        text = _extract_text_from_pdf(pdf_path)
        if text:
            text_source = text
        if has_widget_data:
            result.update(_merge_records(result, widget_data))
            if text:
                text_record = _parse_text_content(text)
                result = _merge_records(result, text_record)
                result["method"] = "mixed"
            else:
                result["method"] = "widgets"
        else:
            if _is_text_scanned(text):
                ocr_text = _extract_ocr_text(pdf_path)
                text_source = ocr_text or text
                parsed = _parse_text_content(ocr_text or text)
                result = _merge_records(result, parsed)
                if (result.get("ssn") is None or result.get("ein") is None) and convert_from_path is not None and pytesseract is not None and cv2 is not None and np is not None:
                    images = convert_from_path(pdf_path, dpi=350, first_page=1, last_page=1)
                    if images:
                        image = np.array(images[0])
                        if result.get("ssn") is None:
                            ssn_candidate = _extract_number_near_label(image, [r"social\s+security\s+number", r"\bssn\b"])
                            if ssn_candidate:
                                result["ssn"] = ssn_candidate
                        if result.get("ein") is None:
                            ein_candidate = _extract_number_near_label(image, [r"employer\s+identification\s+number", r"\bein\b", r"\btin\b"])
                            if ein_candidate:
                                result["ein"] = ein_candidate
                result["method"] = "ocr"
            else:
                parsed = _parse_text_content(text)
                result = _merge_records(result, parsed)
                result["method"] = "text"
        if text_source:
            result["raw_snippet"] = text_source[:800]
        result["status"] = _determine_status(result)
    except Exception as exc:
        logger.exception("Failed to extract W9: %s", exc)
        result["status"] = "ERROR"
        result["method"] = result.get("method") or "ocr"
        if not result["raw_snippet"] and text_source:
            result["raw_snippet"] = text_source[:800]
    return result


if __name__ == "__main__":
    import argparse
    import json

    parser = argparse.ArgumentParser(description="Extract W-9 fields from a PDF file.")
    parser.add_argument("pdf", help="Path to the W-9 PDF file")
    args = parser.parse_args()
    record = extract_w9(args.pdf)
    print(json.dumps(record, indent=2, ensure_ascii=False))
