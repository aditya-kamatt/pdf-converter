import os
import subprocess
import tempfile
import shutil
import pytest

FIXTURE_DIR = os.path.join(os.path.dirname(__file__), '../fixtures')
SAMPLE_PDF = os.path.join(FIXTURE_DIR, 'sample1.pdf')
GOLDEN_XLSX = os.path.join(FIXTURE_DIR, 'golden_sample1.xlsx')

@pytest.mark.skipif(not os.path.exists(SAMPLE_PDF), reason='Fixture PDF not found')
def test_cli_end_to_end():
    with tempfile.TemporaryDirectory() as tmpdir:
        out_xlsx = os.path.join(tmpdir, 'out.xlsx')
        result = subprocess.run([
            'python', '-m', 'orders_converter', SAMPLE_PDF, '-o', out_xlsx
        ], capture_output=True, text=True)
        assert result.returncode == 0
        assert os.path.exists(out_xlsx)
        # If golden file exists, compare checksums
        if os.path.exists(GOLDEN_XLSX):
            import hashlib
            def file_md5(path):
                with open(path, 'rb') as f:
                    return hashlib.md5(f.read()).hexdigest()
            assert file_md5(out_xlsx) == file_md5(GOLDEN_XLSX) 