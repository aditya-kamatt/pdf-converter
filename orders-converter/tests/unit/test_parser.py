import os
import pytest
from orders_converter.core import parser

FIXTURE_DIR = os.path.join(os.path.dirname(__file__), '..', 'fixtures')
SAMPLE_PDF = os.path.join(FIXTURE_DIR, 'sample1.pdf')

@pytest.mark.skipif(not os.path.exists(SAMPLE_PDF), reason='Fixture PDF not found')
def test_extract_header_meta():
    meta = parser.extract_header_meta(SAMPLE_PDF)
    assert isinstance(meta, dict)
    assert 'po_number' in meta
    assert meta['page_count'] > 0

@pytest.mark.skipif(not os.path.exists(SAMPLE_PDF), reason='Fixture PDF not found')
def test_extract_table_rows():
    rows = parser.extract_table_rows(SAMPLE_PDF)
    assert isinstance(rows, list)
    assert all(isinstance(row, list) for row in rows)
    assert len(rows) > 0 