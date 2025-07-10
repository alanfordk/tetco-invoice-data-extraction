import fitz               # PyMuPDF
import re
from dateutil import parser as dateparser
from .utils import parse_amount

def _fmt_date(dt):
    return dt.strftime("%m/%d/%Y")

def extract_invoice_data(path):
    # 1) Load text
    doc = fitz.open(path)
    text = "\n".join(page.get_text() or "" for page in doc)

    # 2) Header
    header = re.search(
        r"Invoice\s*(?:Number|No\.?)[:\s]*([\w-]+).*?"
        r"Date[:\s]*([A-Za-z]{3,}\.?\s*\d{1,2},\s*\d{4})",
        text, re.DOTALL|re.IGNORECASE
    )
    if header:
        invoice_number = header.group(1)
        try:
            inv_dt = dateparser.parse(header.group(2)).date()
            invoice_date = _fmt_date(inv_dt)
        except:
            invoice_date = None
    else:
        invoice_number = invoice_date = None

    # 3) Split on item numbers
    blocks = re.split(r"\r?\n(?=\s*\d+\s*\r?\n)", text)

    # 4) Filter to real item blocks
    item_blocks = []
    for b in blocks:
        lines = [ln.strip() for ln in b.splitlines() if ln.strip()]
        if lines and re.fullmatch(r"\d+", lines[0]):
            item_blocks.append(lines)

    records = []
    amt_re1 = re.compile(r"^\$\s?([\d,]+(?:\.\d+)?)$")
    amt_re2 = re.compile(r"^[\d,]+(?:\.\d+)?$")

    for lines in item_blocks:
        item_idx = int(lines[0])

        # 5a) Amount
        invoice_amount = None
        amt_idx = len(lines)
        for idx, ln in enumerate(lines):
            ln_str = ln.strip()
            if m1 := amt_re1.match(ln_str):
                invoice_amount = parse_amount(m1.group(1)); amt_idx = idx; break
            if ln_str == "$" and idx+1 < len(lines):
                if m2 := amt_re2.match(lines[idx+1].strip()):
                    invoice_amount = parse_amount(m2.group(0)); amt_idx = idx; break
            if idx and (m3 := amt_re2.match(ln_str)):
                invoice_amount = parse_amount(m3.group(0)); amt_idx = idx; break

        # 5b) Description
        description = " ".join(lines[1:amt_idx]).strip()
        description = f"{description}\nItem {item_idx}"

        # 5c) Date ranges
        date_ranges = []

        # (i) Cross‐month hyphen
        for mo1,d1,mo2,d2,yr in re.findall(
            r"([A-Za-z]+)\s+(\d{1,2})-(?:\s*)([A-Za-z]+)\s+(\d{1,2}),\s*(\d{4})",
            description
        ):
            try:
                s = dateparser.parse(f"{mo1} {d1}, {yr}").date()
                e = dateparser.parse(f"{mo2} {d2}, {yr}").date()
                date_ranges.append((s,e))
            except: pass

        # (ii) Same‐month ranges
        for mo,d1,d2,yr in re.findall(
            r"([A-Za-z]+)\s+(\d{1,2})-(\d{1,2}),\s*(\d{4})",
            description
        ):
            try:
                s = dateparser.parse(f"{mo} {d1}, {yr}").date()
                e = dateparser.parse(f"{mo} {d2}, {yr}").date()
                date_ranges.append((s,e))
            except: pass

        # (iii) Fallback “to” / comma‐list
        if len(date_ranges) <= 1:
            if m := re.search(
                r"Project\s+Dates?:\s*(.*?),\s*(\d{4})",
                description, re.IGNORECASE
            ):
                raw, yr = m.groups()
                if m_to := re.match(
                    r"([A-Za-z]+)\s+(\d{1,2})\s+to\s+(?:([A-Za-z]+)\s+)?(\d{1,2})",
                    raw, re.IGNORECASE
                ):
                    mo1,d1,mo2,d2 = m_to.groups(); mo2=mo2 or mo1
                    try:
                        s = dateparser.parse(f"{mo1} {d1}, {yr}").date()
                        e = dateparser.parse(f"{mo2} {d2}, {yr}").date()
                        date_ranges = [(s,e)]
                    except: date_ranges=[]
                else:
                    parts = [p.strip() for p in raw.split(",")]
                    prev=None
                    for part in parts:
                        if " " in part:
                            prev_month,days = part.split(" ",1)
                        else:
                            days = part; prev_month = prev_month
                        if "-" in days:
                            d1,d2 = days.split("-",1)
                        else:
                            d1=d2=days
                        try:
                            s = dateparser.parse(f"{prev_month} {d1}, {yr}").date()
                            e = dateparser.parse(f"{prev_month} {d2}, {yr}").date()
                            date_ranges.append((s,e))
                        except: pass

        # (iv) “Test date:” fallback
        if not date_ranges:
            if m_s := re.search(
                r"Test\s*date[s]?:\s*([A-Za-z]+\s+\d{1,2},\s*\d{4})",
                description, re.IGNORECASE
            ):
                try:
                    dt = dateparser.parse(m_s.group(1)).date()
                    date_ranges.append((dt,dt))
                except: pass

        # ✱ FINAL fallback: any standalone “Month D, YYYY” ✱
        if not date_ranges:
            if m_any := re.search(
                r"([A-Za-z]+\s+\d{1,2},\s*\d{4})",
                description
            ):
                try:
                    dt = dateparser.parse(m_any.group(1)).date()
                    date_ranges.append((dt,dt))
                except: pass

        # 5d) Collapse to overall dates
        if date_ranges:
            test_date_start = _fmt_date(min(s for s,_ in date_ranges))
            test_date_end   = _fmt_date(max(e for _,e in date_ranges))
        else:
            test_date_start = test_date_end = None

        # 5e) Build record
        records.append({
            "client_name":         "client_name",
            "location":            "location",
            "invoice_number":      invoice_number,
            "invoice_date":        invoice_date,
            "project_description": description,
            "invoice_amount":      invoice_amount,
            "test_date_start":     test_date_start,
            "test_date_end":       test_date_end,
        })

    # 6) Invoice‐level mobilization_count
    extra = sum(
        1 for rec in records
        if re.search(r"\bmobilization\b", rec["project_description"], re.IGNORECASE)
    )
    count = 1 + extra
    for rec in records:
        rec["mobilization_count"] = count

    return records
