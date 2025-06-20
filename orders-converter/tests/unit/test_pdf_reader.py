import os
import pytest
from orders_converter.io.pdf_reader import read_pdf_table_and_meta

FIXTURE_DIR = os.path.join(os.path.dirname(__file__), '../fixtures')
SAMPLE_PDF = os.path.join(FIXTURE_DIR, 'sample1.pdf')

@pytest.mark.skipif(not os.path.exists(SAMPLE_PDF), reason='Fixture PDF not found')
def test_read_pdf_table_and_meta():
    meta, rows = read_pdf_table_and_meta(SAMPLE_PDF)
    assert isinstance(meta, dict)
    assert isinstance(rows, list) 