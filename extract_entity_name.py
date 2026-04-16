"""
Simple extractor for extracting the Entity Name from a W-9 form.
"""

import re
import unicodedata
import logging
from pathlib import Path
from typing import Optional

logging.basicConfig(level=logging.WARNING, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

# OCR whitelist for Tesseract
OCR_WHITELIST = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz.,-/() "
OCR_CONFIG = f'--oem 3 --psm 6 -c tessedit_char_whitelist="{OCR_WHITELIST}"'
SSN_PATTERN = re.compile(r"\b(\d{3})[-\s]*(\d{2})[-\s]*(\d{4})\b")

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    from pdf2image import convert_from_path
except ImportError:
    convert_from_path = None

try:
    import pytesseract
except ImportError:
    pytesseract = None

try:
    import cv2
except ImportError:
    cv2 = None

try:
    import numpy as np
except ImportError:
    np = None


# Patterns for determining business type (TOB) from checkboxes
CHECKBOX_PATTERNS = {
    "Individual/sole proprietor": [r"individual/?sole proprietor", r"individual sole proprietor"],
    "C corporation": [r"c(?:\s|-)corporation", r"c corporation", r"c corp"],
    "S corporation": [r"s(?:\s|-)corporation", r"s corporation", r"s corp"],
    "Partnership": [r"partnership"],
    "Trust/estate": [r"trust/?estate", r"trust estate"],
    "Other": [r"\bother\b"],
}

DATE_PATTERN = re.compile(r"(?P<m>\d{1,2})[\./-](?P<d>\d{1,2})[\./-](?P<y>\d{2,4})")


def _normalize_text(value: Optional[str]) -> str:
    """Normalize text for processing."""
    if value is None:
        return ""
    if isinstance(value, bytes):
        value = value.decode("utf-8", errors="ignore")
    text = str(value)
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\r", "\n")
    text = re.sub(r"[\t\x0b\x0c]+", " ", text)
    return text.strip()


def _digits_before_closing_parenthesis(text: str) -> str:
    """Extract digits that appear immediately before a closing parenthesis.

    This helps recover SSN digits from noisy OCR output like "1) 1) 171) -])3)3)-]) 3) 3) 3) 3".
    """
    return "".join(re.findall(r"(\d)(?=\))", text))


def _is_text_scanned(text: str) -> bool:
    """Determine if text appears scanned (few words)."""
    if not text:
        return True
    words = re.findall(r"\w+", text)
    return len(words) < 40


def _preprocess_ocr_image(image: "np.ndarray") -> Optional["np.ndarray"]:
    """Preprocess images for OCR quality."""
    if cv2 is None or np is None:
        return None
    
    try:
        gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
        blurred = cv2.medianBlur(gray, 3)
        _, thresh = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        return thresh
    except Exception as e:
        logger.error(f"Image preprocessing failed: {e}")
        return None


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

                if ")" in combined:
                    digits = _digits_before_closing_parenthesis(combined)
                else:
                    digits = re.sub(r"[^0-9]", "", combined)

                if len(digits) >= required_digits:
                    return digits[:required_digits]
                break
    return None


def _extract_ssn_from_boxed_area(image: "np.ndarray") -> Optional[str]:
    if pytesseract is None:
        return None

    words = _get_ocr_words(image)
    if not words:
        return None

    lines = _group_words_by_line(words, y_tolerance=15)
    best_candidate = None
    best_score = 0

    page_width = image.shape[1]
    for _, line_words in lines:
        if not line_words:
            continue
        line_text = " ".join(w["text"] for w in line_words)
        if ")" in line_text:
            digits = _digits_before_closing_parenthesis(line_text)
        else:
            digits = re.sub(r"[^0-9]", "", line_text)
        if len(digits) < 9:
            continue

        numeric_tokens = sum(1 for w in line_words if re.search(r"[0-9]", w["text"]))
        alpha_tokens = sum(1 for w in line_words if re.search(r"[A-Za-z]", w["text"]))
        avg_left = sum(w["left"] for w in line_words) / len(line_words)
        right_bias = avg_left / max(1, page_width)

        score = len(digits) * 10 + numeric_tokens * 5 - alpha_tokens * 3 + int(right_bias * 10)
        if score > best_score:
            best_score = score
            best_candidate = digits

    if best_candidate and len(best_candidate) >= 9:
        return f"{best_candidate[:3]}-{best_candidate[3:5]}-{best_candidate[5:9]}"

    return None


def _is_label_remainder(value: str) -> bool:
    """Check whether a string is a label remainder or instruction."""
    if not value:
        return False
    remainder = value.lower().strip()
    remainder = re.sub(r"^[\W_]+", "", remainder)
    
    # Special handling for slashes in labels
    if remainder in ("individual", "sole proprietor", "corporation", "partnership", "entity", "indiv"):
        return True
    
    return bool(
        re.search(
            r"^(an entry is required|for a sole proprietor|if different from above|line \d+|entitys? name|business name|address|number, street|apt\.|suite no|city, state|city.*state|state.*zip|and zip code|zip code|see instructions|enter the|requester|purpose of form|disregarded entity|if different|name is required|income tax return|show on|as shown on|not leave)",
            remainder,
        )
    )


def _find_label_value(lines: list, label_patterns: list) -> Optional[str]:
    """Find a value after a label in a list of lines.
    
    Works for multiple formats:
    - Label: Value (on one line)
    - Label
      Value (on the next line)
    """
    for index, line in enumerate(lines):
        normalized = line.strip()
        
        for pattern in label_patterns:
            match = re.search(pattern, normalized, re.IGNORECASE)
            if not match:
                continue
            
            # 1. Check the text after the label on the same line
            after = normalized[match.end():].strip(" :.-/_\\()")
            if after:
                after = re.sub(r"^[\./\-\s]+", "", after)
                
                # Remove trailing dots and instruction words  
                after = re.sub(r'\s+(Name|is required|do not leave|See instructions|not different).*$', '', after, flags=re.IGNORECASE)
            
            if after and len(after) > 1 and not _is_label_remainder(after):
                return after
            
            # 2. If nothing useful follows the label, search the next lines
            for next_line in lines[index + 1 : index + 7]:
                if not next_line:
                    continue
                
                candidate = next_line.strip()
                candidate = re.sub(r"^[\./\-\s]+", "", candidate)
                
                # Skip if this is clearly another label or instruction
                if re.search(
                    r"^(\d+\s+|Name|Business|Address|City|State|ZIP|SSN|EIN|Date|Signature|For|See|An entry|If different)", 
                    candidate,
                    re.IGNORECASE
                ):
                    continue
                
                if candidate and len(candidate) > 1 and not _is_label_remainder(candidate):
                    return candidate
    
    return None


def _extract_ssn_from_text(text: str) -> Optional[str]:
    """Extract a Social Security Number from text."""
    if not text:
        return None
    candidate = _normalize_text(text)
    match = SSN_PATTERN.search(candidate)
    if match:
        return f"{match.group(1)}-{match.group(2)}-{match.group(3)}"
    fallback = re.search(r"\b(\d{3})\s+(\d{2})\s+(\d{4})\b", candidate)
    if fallback:
        return f"{fallback.group(1)}-{fallback.group(2)}-{fallback.group(3)}"
    return _extract_ssn_from_nearby_label_text(candidate)


def _extract_ssn_from_nearby_label_text(text: str) -> Optional[str]:
    if not text:
        return None

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return None

    label_pattern = re.compile(r"social\s+security\s+number", re.IGNORECASE)

    for index, line in enumerate(lines):
        if label_pattern.search(line):
            digits = ""
            for next_index in range(index, min(index + 4, len(lines))):
                current_line = lines[next_index]
                if ")" in current_line:
                    digits += _digits_before_closing_parenthesis(current_line)
                else:
                    digits += re.sub(r"[^0-9]", "", current_line)

                if len(digits) >= 9:
                    return f"{digits[:3]}-{digits[3:5]}-{digits[5:9]}"
            return None

    return None


def _extract_name_entity_from_text(text: str) -> Optional[str]:
    """Extract the Entity Name from W-9 text."""
    if not text:
        return None
    
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    # Specialized search for W-9 (March 2024 and later)
    # Find a line starting with "Name (" containing instructions
    for index, line in enumerate(lines):
        # Pattern: "Name (as shown on your income tax return)"
        if re.search(r"^Name\s*\(.+\)", line, re.IGNORECASE):
            # The value should be on the next line
            if index + 1 < len(lines):
                next_line = lines[index + 1].strip()
                if next_line and not _is_label_remainder(next_line) and len(next_line) > 2:
                    return next_line
    
    # Search for the pattern "1 Name of entity/individual"
    for index, line in enumerate(lines):
        if re.search(r"^1\s+Name of entity", line, re.IGNORECASE):
            # Skip instructions on the same line
            # Search the next 10 lines for the actual name
            for next_idx in range(index + 1, min(index + 10, len(lines))):
                candidate = lines[next_idx].strip()
                
                # Skip blank lines and instructions
                if not candidate or len(candidate) < 3:
                    continue
                
                # Skip if this is an instruction or another label
                if re.search(r"^(An entry|For a|If|See|Business|Name|Address|City|\d+\s)", candidate, re.IGNORECASE):
                    continue
                
                # This is most likely the name
                if _is_label_remainder(candidate):
                    continue
                return candidate
    
    # If primary patterns fail, try a generic label search
    result = _find_label_value(
        lines,
        [
            r"^Name of entity/individual",
            r"^Name of entity",
        ],
    )
    
    if result:
        return result
    return None


def _extract_tob(text: str) -> Optional[str]:
    """Extract business type (TOB) from a W-9 checkbox field (field 3a).

    The selection indicator is often OCR text immediately before the checkbox label.
    Examples:
    - 'etor C corporation' means C corporation is selected
    - 'hip Trust/estate' means Trust/estate is selected
    - if no checkbox letters are found, search for "Other" on a separate line
    """
    if not text:
        return None
    
    candidate = _normalize_text(text)
    lines = candidate.split('\n')
    checkbox_lines = []
    
    # Find all lines containing business types
    for i, line in enumerate(lines):
        line_lower = line.lower()
        types_count = sum([1 for t in ['individual', 'corporation', 'partnership', 'trust', 'other'] 
                          if t in line_lower])
        # Add lines that contain:
        # - multiple types (main checkboxes), OR
        # - the word "other" with "see" (separate line for Other)
        if types_count > 2 or ('other' in line_lower and 'see' in line_lower):
            checkbox_lines.append((i, line))
    
    if not checkbox_lines:
        return None
    
    # Types to check
    types_to_check = [
        ("C Corporation", r'c\s*corporation'),
        ("S Corporation", r's\s*corporation'),
        ("Partnership", r'partnership'),
        ("Trust/estate", r'trust\s*[/\s]*estate'),
        ("Individual/sole proprietor", r'individual\s+(?:sole\s+)?proprietor'),
    ]
    
    # Known OCR separators (usually before unchecked boxes)
    ocr_delimiters = {'i', 'im', 'el', 'cj', 'j', 'tc', 'gs', 'e', 'z', 'c'}
    
    main_line = checkbox_lines[0][1].lower()
    matches_with_letters = []  # Types with letters before them (excluding delimiters)
    
    for tob_name, pattern in types_to_check:
        match = re.search(pattern, main_line)
        if not match:
            continue
        
        start_pos = match.start()
        
        # Take a few characters before the type
        if start_pos == 0:
            before = ""
        else:
            # Take up to 4 characters before the type
            before = main_line[max(0, start_pos-4):start_pos]
        
        # Strip whitespace
        before_clean = before.strip()
        
        # KEY TEST: are there only letters before it and is it not a known delimiter?
        is_only_letters_and_not_delimiter = (before_clean and before_clean.isalpha() 
                                             and before_clean.lower() not in ocr_delimiters)
        # If only letters appear before the type (and it is not a delimiter), this is a selected candidate
        if is_only_letters_and_not_delimiter:
            matches_with_letters.append((tob_name, len(before_clean)))
    
    # If types with letters were found, return the best candidate
    if matches_with_letters:
        # Sort by length before the type (longer is more likely real letters)
        matches_with_letters.sort(key=lambda x: -x[1])
        best_tob = matches_with_letters[0][0]
        return best_tob
    
    # If the main line has no letters before the type,
    # search for "Other" on a separate line
    for idx, line in checkbox_lines:
        if 'other' in line.lower() and 'see' in line.lower():
            return "Other"
    return None


def _clean_business_name(name: Optional[str]) -> Optional[str]:
    """Clean Business Name from OCR artifacts."""
    if not name:
        return name
    
    # Remove leading punctuation and brackets
    cleaned = re.sub(r"^[\.,\-\(\)\[\]\s]+", "", name).strip()
    
    # Remove trailing punctuation
    cleaned = re.sub(r"[\.,\-\(\)\[\]\s]+$", "", cleaned).strip()
    
    return cleaned if cleaned else None


def _extract_business_name_from_text(text: str) -> Optional[str]:
    """Extract Business Name from W-9 text."""
    if not text:
        return None
    
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    # Search for the Business Name pattern (field 2)
    for index, line in enumerate(lines):
        # Pattern: "2 Business name"
        if re.search(r"^2\s+Business name", line, re.IGNORECASE):
            # Search for the value in following lines
            for next_idx in range(index + 1, min(index + 8, len(lines))):
                candidate = lines[next_idx].strip()
                
                if not candidate or len(candidate) < 2:
                    continue
                
                # Skip instructions and labels
                if re.search(r"^(If|See|Address|Name|Type|City|\d+\s)", candidate, re.IGNORECASE):
                    continue
                
                if _is_label_remainder(candidate):
                    continue
                
                # Clean artifacts
                cleaned = _clean_business_name(candidate)
                return cleaned
    
    # Alternative search
    result = _find_label_value(
        lines,
        [
            r"^2\s+Business name",
            r"^Business name",
            r"disregarded entity name",
        ],
    )
    
    if result:
        cleaned = _clean_business_name(result)
        return cleaned
    return None


def _extract_address_from_text(text: str) -> Optional[str]:
    """Extract Address from W-9 text (field 5)."""
    if not text:
        return None
    
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    # Search for the Address pattern (field 5): "5 Address (number, street..."
    for index, line in enumerate(lines):
        # Pattern starting with digit 5 and the word Address
        if re.search(r"^[35]\s+.*[Aa]ddress\s*\(", line, re.IGNORECASE):
            # Search for the value in following lines (maximum 3-4 lines)
            for next_idx in range(index + 1, min(index + 5, len(lines))):
                candidate = lines[next_idx].strip()
                
                if not candidate or len(candidate) < 3:
                    continue
                
                # Skip if this is clearly another label or instruction
                if re.search(r"^(City|State|ZIP|6\s+|See instructions|Requesters name)", candidate, re.IGNORECASE):
                    continue
                
                # This is most likely the address
                return candidate
    return None


def _extract_city_state_zip_from_text(text: str) -> Optional[str]:
    """Extract City, State, and ZIP from W-9 text (field 6).
    
    Works with OCR variants such as:
    - "6 City, state, and ZIP code"
    - "6 Clty, state, and ZIP code" (OCR typo)
    - "6 City state ZIP" (different formats)
    """
    if not text:
        return None
    
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    # Search for the City/State/ZIP pattern (field 6): "6 ..." containing state and zip
    # More flexible pattern for OCR errors such as "Clty" instead of "City"
    for index, line in enumerate(lines):
        # Look for a line starting with "6 " containing "state" and "zip"
        if re.search(r"^6\s+", line, re.IGNORECASE) and \
           (re.search(r"state", line, re.IGNORECASE) and re.search(r"zip", line, re.IGNORECASE)):
            # Search for the value in following lines (maximum 3-4 lines)
            for next_idx in range(index + 1, min(index + 5, len(lines))):
                candidate = lines[next_idx].strip()
                
                if not candidate or len(candidate) < 5:  # Minimum "City, ST ZIP"
                    continue
                
                # Skip if this is clearly another label or instruction
                if re.search(r"^(List account|7\s+|Requesters|See instructions)", candidate, re.IGNORECASE):
                    continue
                
                # This is most likely City/State/ZIP
                return candidate
    return None


def _extract_from_widgets(pdf_path: str) -> Optional[tuple]:
    """Attempt to extract fields from PDF widgets."""
    if fitz is None:
        return None
    
    try:
        doc = fitz.open(pdf_path)
        entity_name = None
        business_name = None
        tob = None
        address = None
        city_state_zip = None
        ssn = None
        ssn_parts = {}
        
        # Map checkbox values c1_1 to business types (1-indexed for W-9)
        # The field value is the 1-indexed position in the type list (not "Off")
        # W-9 Section 3a order: Individual, C Corp (2nd), S Corp, Partnership (4th), Trust/estate, Other
        tob_mapping = {
            '1': "Individual/sole proprietor",
            '2': "C Corporation",          # ← "C Corporation second option (2nd item)"
            '3': "S Corporation",
            '4': "Partnership",            # ← "Partnership fourth option (4th item)"
            '5': "Trust/estate",
            '6': "Other",
        }
        
        for page_idx, page in enumerate(doc):
            widgets = list(page.widgets())
            if not widgets:
                continue

            for widget in widgets:
                field_name = widget.field_name or ""
                value = str(widget.field_value or "").strip()
                if not value:
                    continue

                # W-9: f1_01 = Entity Name
                if "f1_01" in field_name:
                    if len(value) > 1:
                        entity_name = value
                        continue

                # W-9: f1_02 = Business Name
                if "f1_02" in field_name:
                    if len(value) > 1:
                        business_name = value
                        continue

                # W-9: f1_11, f1_12, f1_13 = SSN parts
                if "f1_11" in field_name:
                    ssn_parts["part1"] = value
                    continue
                if "f1_12" in field_name:
                    ssn_parts["part2"] = value
                    continue
                if "f1_13" in field_name:
                    ssn_parts["part3"] = value
                    continue

                if "ssn" in field_name.lower() and ssn is None:
                    ssn = _extract_ssn_from_text(value)
                    continue

                # TOB checkboxes: c1_1[0] to c1_1[5]
                # Field value indicates a 1-indexed position in the type list (not "Off")
                if "c1_1[" in field_name and tob is None:
                    checkbox_value = value.strip()
                    if checkbox_value and checkbox_value != "Off":
                        if checkbox_value in tob_mapping:
                            tob = tob_mapping[checkbox_value]
                        continue

                # W-9: f1_07 = Address (street and number)
                if "f1_07" in field_name or ("f1_05" in field_name and not address and "Address_ReadOrder" not in field_name):
                    if len(value) > 1:
                        address = value
                        continue

                # W-9: f1_08 = City, State, ZIP code
                if "f1_08" in field_name or ("f1_06" in field_name and not city_state_zip and "Address_ReadOrder" not in field_name):
                    if len(value) > 1:
                        city_state_zip = value
                        continue
        doc.close()

        if ssn is None and ssn_parts.get("part1") and ssn_parts.get("part2") and ssn_parts.get("part3"):
            ssn = f"{ssn_parts['part1']}-{ssn_parts['part2']}-{ssn_parts['part3']}"

        if entity_name or business_name or tob or address or city_state_zip or ssn:
            return (entity_name, business_name, tob, address, city_state_zip, ssn)
    except Exception as e:
        logger.error(f"Widget extraction failed: {e}", exc_info=True)
    
    return None


def _extract_from_text_layer(pdf_path: str) -> Optional[tuple]:
    """Extract fields from the PDF text layer."""
    if fitz is None:
        return None

    try:
        doc = fitz.open(pdf_path)
        full_text = ""
        for page_idx, page in enumerate(doc):
            page_text = page.get_text()
            full_text += page_text
        doc.close()

        if not full_text or not full_text.strip():
            return None

        entity_name = _extract_name_entity_from_text(full_text)
        business_name = _extract_business_name_from_text(full_text)
        tob = _extract_tob(full_text)
        address = _extract_address_from_text(full_text)
        city_state_zip = _extract_city_state_zip_from_text(full_text)
        ssn = _extract_ssn_from_text(full_text)

        if entity_name or business_name or tob or address or city_state_zip or ssn:
            return (entity_name, business_name, tob, address, city_state_zip, ssn)
    except Exception as e:
        logger.error(f"Text layer extraction failed: {e}", exc_info=True)

    return None


def _render_pdf_page_to_image(pdf_path: str, page_number: int = 0, dpi: int = 350) -> Optional["np.ndarray"]:
    """Render a PDF page to an RGB image using PyMuPDF."""
    if fitz is None or np is None:
        return None

    try:
        with fitz.open(pdf_path) as doc:
            if page_number >= doc.page_count:
                return None
            page = doc.load_page(page_number)
            matrix = fitz.Matrix(dpi / 72.0, dpi / 72.0)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            shape = (pix.height, pix.width, pix.n)
            image = np.frombuffer(pix.samples, dtype=np.uint8).reshape(shape)
            if pix.n == 1:
                image = np.stack([image] * 3, axis=-1)
            elif pix.n == 4:
                image = image[:, :, :3]
            return image
    except Exception as e:
        logger.debug(f"PDF rendering failed: {e}")
        return None


def _extract_from_ocr(pdf_path: str) -> Optional[tuple]:
    """Use OCR to extract fields from scanned PDF."""
    if pytesseract is None:
        logger.warning("pytesseract is missing, OCR is unavailable")
        return None

    images = None
    if convert_from_path is not None:
        try:
            images = convert_from_path(pdf_path, dpi=350, first_page=1, last_page=1)
        except Exception as e:
            logger.warning(f"pdf2image conversion failed: {e}")

    if not images:
        image = _render_pdf_page_to_image(pdf_path, dpi=350)
        if image is None:
            logger.warning("Unable to render PDF page for OCR")
            return None
    else:
        image = np.array(images[0]) if np is not None else images[0]

    processed = _preprocess_ocr_image(image)
    if processed is None:
        processed = image

    try:
        text = pytesseract.image_to_string(processed, config=OCR_CONFIG)
        if not text.strip() or len(text.strip()) < 10:
            if cv2 is not None:
                inverted = cv2.bitwise_not(processed)
                text = pytesseract.image_to_string(inverted, config=OCR_CONFIG)

        if text:
            entity_name = _extract_name_entity_from_text(text)
            business_name = _extract_business_name_from_text(text)
            tob = _extract_tob(text)
            ssn = _extract_ssn_from_text(text)

            if ssn is None and image is not None:
                ssn = _extract_ssn_from_boxed_area(image)
            if ssn is None and image is not None:
                ssn_digits = _extract_number_near_label(
                    image,
                    [r"social\s+security\s+number"],
                    required_digits=9,
                )
                if ssn_digits and len(ssn_digits) == 9:
                    ssn = f"{ssn_digits[:3]}-{ssn_digits[3:5]}-{ssn_digits[5:]}"

            if entity_name or business_name or tob or ssn:
                address = _extract_address_from_text(text)
                city_state_zip = _extract_city_state_zip_from_text(text)
                return (entity_name, business_name, tob, address, city_state_zip, ssn)
        else:
            logger.warning("Tesseract returned no text after both attempts")
    except Exception as e:
        logger.error(f"OCR extraction failed: {e}", exc_info=True)

    return None


def extract_entity_name(pdf_path: str) -> dict:
    """Main function to extract W-9 data from a PDF."""
    pdf_path = str(pdf_path)
    result = {
        "entity_name": None,
        "business_name": None,
        "tob": None,
        "address": None,
        "city_state_zip": None,
        "ssn": None,
        "method": None,
        "success": False,
    }

    if not Path(pdf_path).exists():
        logger.error(f"File not found: {pdf_path}")
        return result

    _ = Path(pdf_path).stat().st_size
    method_1 = _extract_from_widgets(pdf_path)
    if method_1:
        entity, business, tob, address, city_state_zip, ssn = method_1
        result.update({
            "entity_name": entity,
            "business_name": business,
            "tob": tob,
            "address": address,
            "city_state_zip": city_state_zip,
            "ssn": ssn,
            "method": "widgets",
            "success": True,
        })
        return result

    method_2 = _extract_from_text_layer(pdf_path)
    if method_2:
        entity, business, tob, address, city_state_zip, ssn = method_2
        result.update({
            "entity_name": entity,
            "business_name": business,
            "tob": tob,
            "address": address,
            "city_state_zip": city_state_zip,
            "ssn": ssn,
            "method": "text",
            "success": True,
        })
        return result

    method_3 = _extract_from_ocr(pdf_path)
    if method_3:
        entity, business, tob, address, city_state_zip, ssn = method_3
        result.update({
            "entity_name": entity,
            "business_name": business,
            "tob": tob,
            "address": address,
            "city_state_zip": city_state_zip,
            "ssn": ssn,
            "method": "ocr",
            "success": True,
        })
        return result

    logger.warning("OCR also failed")
    logger.error("FAILED: Unable to extract data from {pdf_path}".format(pdf_path=pdf_path))
    return result


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python extract_entity_name.py <path_to_pdf>")
        sys.exit(1)

    pdf_file = sys.argv[1]
    result = extract_entity_name(pdf_file)

    print("\n" + "="*50)
    print(f"Result for: {pdf_file}")
    print("="*50)
    print(f"Entity Name:  {result['entity_name']}")
    print(f"Business Name: {result['business_name']}")
    print(f"TOB:           {result['tob']}")
    print(f"SSN:           {result.get('ssn')}")
    print(f"Method: {result['method']}")
    print(f"Success: {result['success']}")
    print(f"Address: {result['address']}")
    print(f"City/State/ZIP: {result['city_state_zip']}")
