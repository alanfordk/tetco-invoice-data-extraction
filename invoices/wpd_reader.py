# invoices/wpd_reader.py

import os
import subprocess
import tempfile
from invoices.pdf_reader import extract_invoice_data as extract_pdf

def extract_invoice_data(path):
    """
    Convert a .wpd to PDF via LibreOffice, then parse with pdf_reader.
    Requires 'soffice' on your PATH.
    """
    # Create a temp dir for the converted PDF
    with tempfile.TemporaryDirectory() as tmpdir:
        # LibreOffice headless conversion
        subprocess.run([
            "soffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", tmpdir,
            path
        ], check=True)

        # Construct the converted PDF’s filename
        base = os.path.basename(path)
        name, _ = os.path.splitext(base)
        pdf_path = os.path.join(tmpdir, name + ".pdf")

        # Delegate to your pdf_reader
        records = extract_pdf(pdf_path)

    return records
