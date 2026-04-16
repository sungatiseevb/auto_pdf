"""
Batch process all PDF files in input_pdfs.
Save results to CSV and show summary statistics.
"""

import csv
import logging
from pathlib import Path
from datetime import datetime
from extract_entity_name import extract_entity_name

# Logging configuration
logging.basicConfig(
    level=logging.WARNING,
    format='%(asctime)s [%(levelname)s] %(message)s'
)
logger = logging.getLogger(__name__)

# Parameters
INPUT_DIR = Path("input_pdfs")
OUTPUT_DIR = Path("output_pdfs")
RESULTS_FILE = "extraction_results.csv"

# Create output directory if it does not exist
OUTPUT_DIR.mkdir(exist_ok=True)

def batch_extract_pdfs():
    """Process all PDF files in INPUT_DIR."""
    
    # Find all PDF files
    pdf_files = sorted(INPUT_DIR.glob("*.pdf"))
    
    if not pdf_files:
        logger.error(f"No PDF files found in {INPUT_DIR}")
        return
    
    logger.warning("="*70)
    logger.warning(f"BATCH EXTRACTION: Found {len(pdf_files)} PDF files")
    logger.warning("="*70)
    
    results = []
    statistics = {
        "total": len(pdf_files),
        "success": 0,
        "partial": 0,
        "failed": 0,
        "by_method": {
            "widgets": 0,
            "text": 0,
            "ocr": 0,
            "none": 0,
        }
    }
    
    # Process each PDF
    for idx, pdf_file in enumerate(pdf_files, 1):
        logger.warning(f"\n[{idx}/{len(pdf_files)}] Processing: {pdf_file.name}")
        
        try:
            result = extract_entity_name(str(pdf_file))
            
            # Add the result
            record = {
                "filename": pdf_file.name,
                "entity_name": result.get("entity_name", ""),
                "business_name": result.get("business_name", ""),
                "tob": result.get("tob", ""),
                "address": result.get("address", ""),
                "city_state_zip": result.get("city_state_zip", ""),
                "ssn": result.get("ssn", ""),
                "method": result.get("method", ""),
                "success": result.get("success", False),
                "timestamp": datetime.now().isoformat(),
            }
            results.append(record)
            
            # Update statistics
            if result["success"]:
                statistics["success"] += 1
            else:
                if result.get("entity_name"):
                    statistics["partial"] += 1
                else:
                    statistics["failed"] += 1
            
            method = result.get("method") or "none"
            if method in statistics["by_method"]:
                statistics["by_method"][method] += 1
            
            # Output the result
            if result["success"]:
                logger.warning(f"  ✓ Entity: {result['entity_name']}")
                if result.get('business_name'):
                    logger.info(f"    Business: {result['business_name']}")
                if result.get('address'):
                    logger.warning(f"    Address: {result['address']}")
                if result.get('city_state_zip'):
                    logger.warning(f"    City, State, ZIP: {result['city_state_zip']}")
                logger.warning(f"    Method: {result['method']}")
            else:
                logger.warning(f"  ✗ Extraction failed (method: {result['method']})")
                
        except Exception as e:
            logger.error(f"  ❌ Error processing file: {e}", exc_info=True)
            results.append({
                "filename": pdf_file.name,
                "entity_name": "",
                "business_name": "",
                "tob": "",
                "address": "",
                "city_state_zip": "",
                "ssn": "",
                "method": "error",
                "success": False,
                "timestamp": datetime.now().isoformat(),
            })
            statistics["failed"] += 1
    
    # Save results to CSV
    logger.info(f"\n{'='*70}")
    logger.warning(f"Saving results to {RESULTS_FILE}...")
    
    if results:
        with open(RESULTS_FILE, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=results[0].keys())
            writer.writeheader()
            writer.writerows(results)
        logger.warning(f"✓ Results saved: {RESULTS_FILE}")
    
    # Print statistics
    logger.info(f"\n{'='*70}")
    logger.warning("STATISTICS:")
    logger.warning("="*70)
    logger.warning(f"Total files:        {statistics['total']}")
    logger.warning(f"✓ Successful:           {statistics['success']} ({statistics['success']*100//statistics['total']}%)")
    logger.warning(f"⚠ Partial success:   {statistics['partial']} ({statistics['partial']*100//statistics['total']}%)")
    logger.warning(f"✗ Failed:            {statistics['failed']} ({statistics['failed']*100//statistics['total']}%)")
    logger.warning(f"\nBy method:")
    logger.warning(f"  Widgets:  {statistics['by_method']['widgets']}")
    logger.warning(f"  Text:     {statistics['by_method']['text']}")
    logger.warning(f"  OCR:      {statistics['by_method']['ocr']}")
    logger.warning(f"  Error:   {statistics['by_method']['none']}")
    
    # Results table
    logger.warning(f"\n{'='*140}")
    logger.warning("RESULTS:")
    logger.warning("="*140)
    logger.warning(f"{'No':<3} {'Entity Name':<18} {'Address':<28} {'City, State, ZIP':<24} {'TOB':<16} {'SSN':<12} {'Method':<8} {'Status':<8}")
    logger.info("-"*140)
    
    for idx, result in enumerate(results, 1):
        status = "✓" if result["success"] else "✗"
        entity = (result["entity_name"] or "N/A")[:18]
        address = (result.get("address", "") or "—")[:28]
        city_state_zip = (result.get("city_state_zip", "") or "—")[:23]
        tob = (result.get("tob", "") or "—")[:16]
        ssn = (result.get("ssn", "") or "—")[:11]
        method = (result["method"] or "N/A")[:6]
        logger.warning(f"{idx:<3} {entity:<18} {address:<28} {city_state_zip:<24} {tob:<16} {ssn:<12} {method:<8} {status:<8}")
    
    logger.warning("="*70)
    logger.warning("Batch extraction completed!")


if __name__ == "__main__":
    batch_extract_pdfs()