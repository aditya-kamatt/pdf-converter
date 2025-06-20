import argparse
import os
import pandas as pd
from orders_converter.core import parser
from orders_converter.io.excel_writer import write_order_excel

def main():
    parser_arg = argparse.ArgumentParser(description="Orders Sheet Converter CLI")
    parser_arg.add_argument("pdf", help="Path to the purchase order PDF")
    parser_arg.add_argument("-o", "--output", help="Output Excel file path (.xlsx)")
    args = parser_arg.parse_args()

    meta = parser.extract_header_meta(args.pdf)
    rows = parser.extract_table_rows(args.pdf)
    print("Header Meta:")
    for k, v in meta.items():
        print(f"  {k}: {v}")
    print(f"Extracted {len(rows)} table rows.")

    if not rows:
        print("No table rows found. Exiting.")
        return
    df = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows) > 1 else pd.DataFrame()
    out_path = args.output or os.path.splitext(args.pdf)[0] + ".xlsx"
    write_order_excel(df, meta, out_path)
    print(f"Excel file written to: {out_path}")

if __name__ == "__main__":
    main() 