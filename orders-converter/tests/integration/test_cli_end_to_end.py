import os
import subprocess
import tempfile
import shutil
import pytest
import pandas as pd
from pandas.testing import assert_frame_equal

from orders_converter.core.parser import extract_table_rows

FIXTURE_DIR = os.path.join(os.path.dirname(__file__), '../fixtures')
SAMPLE_PDF = os.path.join(FIXTURE_DIR, 'sample1.pdf')
GOLDEN_XLSX = os.path.join(FIXTURE_DIR, 'golden_sample1.xlsx')

@pytest.mark.skipif(not os.path.exists(SAMPLE_PDF), reason='Fixture PDF not found')
def test_cli_end_to_end():
    with tempfile.TemporaryDirectory() as tmpdir:
        out_xlsx = os.path.join(tmpdir, 'out.xlsx')
        result = subprocess.run([
            'python', '-m', 'orders_converter.cli', SAMPLE_PDF, '--output', out_xlsx
        ], capture_output=True, text=True)
        assert result.returncode == 0, f"CLI Error: {result.stderr}"

        assert os.path.exists(out_xlsx)
        
        # Compare content
        expected_df = pd.read_excel(GOLDEN_XLSX)
        actual_df = pd.read_excel(out_xlsx)
        assert_frame_equal(actual_df, expected_df) 