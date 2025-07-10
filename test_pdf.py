# test_pdf.py

import re
import fitz
from invoices.pdf_reader import extract_invoice_data

if __name__ == "__main__":
    sample = r"C:\Users\akitc\OneDrive\Desktop\Data project\Invoices\Cummins Power Generation\Cummins VA Center Cheyenne July 2017 INVOICE 08141701.pdf"

    # 1) Load the full raw text
    doc = fitz.open(sample)
    text = "\n".join(page.get_text() for page in doc)

    # 2) Print a slice of raw text so we can inspect the header & items
    print("=== RAW TEXT (first 2000 chars) ===\n")
    print(text[:2000])
    print("\n... (truncated) ...\n")

    # 3) Split into blocks using the same regex your reader uses
    pattern = r"\n(?=\d+\s)"
    blocks = re.split(pattern, text)
    print(f"=== Total blocks including preamble: {len(blocks)} (so item_blocks = {len(blocks)-1}) ===\n")

    # 4) Dump each block so you can see why there are 6
    for i, block in enumerate(blocks):
        print(f"\n--- Block {i} ---\n")
        print(block[:500])    # print first 500 chars of each block
        print("\n[...]\n")

    # 5) Finally, show what your extractor returns
    print("\n=== EXTRACTED RECORDS ===\n")
    records = extract_invoice_data(sample)
    for i, rec in enumerate(records, start=1):
        print(f"\n=== Item #{i} ===")
        for k, v in rec.items():
            print(f"{k:20}: {v}")
