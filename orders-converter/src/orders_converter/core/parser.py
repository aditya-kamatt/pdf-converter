import pdfplumber
import re
from typing import List, Dict, Any, Tuple


def extract_header_meta(pdf_path: str) -> Dict[str, Any]:
    """
    Extracts header metadata from the purchase order PDF.
    Returns a dictionary with keys: po_number, vendor_number, ship_by_date, payment_terms, total, page_count.
    """
    meta = {
        'po_number': None,
        'vendor_number': None,
        'ship_by_date': None,
        'payment_terms': None,
        'total': None,
        'page_count': 0
    }
    with pdfplumber.open(pdf_path) as pdf:
        meta['page_count'] = len(pdf.pages)
        first_page = pdf.pages[0]
        text = first_page.extract_text() or ""
        # Example regex patterns (adjust as per actual PDF layout)
        patterns = {
            'po_number': r"PO\s*#\s*([A-Za-z0-9-]+)",
            'vendor_number': r"Vendor\s*#\s*([A-Za-z0-9-]+)",
            'ship_by_date': r"Ship[- ]by\s*([\d/\-]+)",
            'payment_terms': r"Payment\s*Terms:\s*([A-Za-z0-9 ]+)",
            'total': r"Total\s*\$([\d,]+\.\d{2})"
        }
        for key, pat in patterns.items():
            m = re.search(pat, text)
            if m:
                meta[key] = m.group(1).strip()
    return meta


def extract_table_rows(pdf_path: str) -> List[List[str]]:
    """
    Extracts table rows from the purchase order PDF using pdfplumber.
    Falls back to regex if table extraction fails.
    Returns a list of rows (each row is a list of cell strings).
    """
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                # Skip header row if repeated on every page
                if rows and table[0] == rows[0]:
                    rows.extend(table[1:])
                else:
                    rows.extend(table)
            else:
                # Fallback: try to extract lines and parse with regex
                text = page.extract_text() or ""
                for line in text.splitlines():
                    # Example: split by 2+ spaces (adjust as needed)
                    cells = re.split(r"\s{2,}", line.strip())
                    if len(cells) > 2:
                        rows.append(cells)
    return rows 