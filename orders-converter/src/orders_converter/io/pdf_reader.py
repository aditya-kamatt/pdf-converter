"""
PDF reading utilities (placeholder).
"""

from orders_converter.core import parser
from typing import Tuple, List, Dict, Any

def read_pdf_table_and_meta(pdf_path: str) -> Tuple[Dict[str, Any], List[List[str]]]:
    """
    Reads the PDF and returns (meta, table_rows).
    """
    meta = parser.extract_header_meta(pdf_path)
    rows = parser.extract_table_rows(pdf_path)
    return meta, rows 