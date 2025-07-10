# test_wpd.py

import os
import subprocess
import tempfile
import fitz
from invoices.wpd_reader import extract_invoice_data

if __name__ == "__main__":
    sample = r"C:\Users\akitc\OneDrive\Desktop\Data project\Invoices\Compass Minerals\2013\D1 Dryer May 2013 INVOICE 05211304.wpd"

    # 1) Convert the .wpd to PDF in a temp dir
    with tempfile.TemporaryDirectory() as tmpdir:
        subprocess.run([
            "soffice", "--headless", "--convert-to", "pdf",
            "--outdir", tmpdir, sample
        ], check=True)

        base = os.path.splitext(os.path.basename(sample))[0]
        pdf_path = os.path.join(tmpdir, base + ".pdf")

        # 2) Open and immediately close the PDF via a context manager
        with fitz.open(pdf_path) as doc:
            raw_text = "\n".join(page.get_text() or "" for page in doc)
        # doc is now closed, so the file is free for deletion

        print("=== RAW TEXT (first 2000 chars) ===\n")
        print(raw_text[:2000])
        print("\n... (truncated) ...\n")

    # 3) Run your extractor on the original .wpd (it will reconvert internally)
    records = extract_invoice_data(sample)

    # 4) Pretty‐print results
    for i, rec in enumerate(records, start=1):
        print(f"\n=== Item #{i} ===")
        for k, v in rec.items():
            print(f"{k:20}: {v}")
