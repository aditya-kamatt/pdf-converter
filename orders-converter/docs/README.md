# Orders Sheet Converter

Convert structured purchase-order PDFs into well-formatted Excel files.

## Quick Start

1. **Install dependencies**
   ```sh
   pip install -r requirements.txt
   # or, if using poetry
   poetry install
   ```

2. **Run the CLI**
   ```sh
   python -m orders_converter <path-to-pdf>
   ```

3. **Run tests**
   ```sh
   pytest
   ```

## Developer Guide

- Source code: `src/orders_converter/`
- Core parsing logic: `core/parser.py`
- Add fixture PDFs to `tests/fixtures/` for real tests.
- To build the GUI, see `gui.py` (to be implemented).
- To package as an EXE: use PyInstaller (see instructions in future docs).

## Project Structure

See the project tree in the main documentation.

## Requirements
- Python 3.8–3.11
- pdfplumber
- pandas
- openpyxl
- pytest

## License
MIT 