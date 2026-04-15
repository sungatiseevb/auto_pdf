import argparse
import json
from pathlib import Path

from w9_extractor_fixed import extract_w9


def print_record(record: dict) -> None:
    print(json.dumps(record, indent=2, ensure_ascii=False))


def load_pdf_files(input_dir: str):
    input_path = Path(input_dir)
    if not input_path.exists():
        print(f"Directory not found: {input_dir}")
        return []

    pdf_files = sorted(input_path.glob("*.pdf"))
    if not pdf_files:
        print(f"No PDFs found in {input_dir}/. Place a PDF there and rerun this script.")
        return []

    return pdf_files


def run_pdf_extraction(input_dir: str = "input_pdfs", output_file: str | None = None) -> None:
    pdf_files = load_pdf_files(input_dir)
    if not pdf_files:
        return

    results = []
    print(f"\n=== PDF EXTRACTION FROM {input_dir} ===")

    for pdf_path in pdf_files:
        print(f"\nFile: {pdf_path.name}")
        record = extract_w9(str(pdf_path))
        print_record(record)
        results.append(record)

    if output_file:
        output_path = Path(output_file)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with output_path.open("w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"\nSaved extraction results to {output_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract W-9 data from PDFs in input_pdfs and optionally save JSON output.")
    parser.add_argument("--input", default="input_pdfs", help="Folder with PDF files")
    parser.add_argument("--output", default=None, help="Optional JSON output file path")
    args = parser.parse_args()

    run_pdf_extraction(args.input, args.output)
