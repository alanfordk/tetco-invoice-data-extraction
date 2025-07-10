import re
from dateutil import parser

def extract_field(pattern, text):
    """Return the first regex group match, or None."""
    m = re.search(pattern, text)
    return m.group(1) if m else None

def parse_date(s):
    """Turn a variety of date strings into a date object or ISO string."""
    if not s: return None
    try:
        return parser.parse(s).date().isoformat()
    except Exception:
        return None

def parse_amount(s):
    """Convert things like '1,234.56' to float."""
    if not s: return None
    return float(s.replace(',', ''))
