import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, numbers
from openpyxl import load_workbook
import logging

def write_to_excel(df: pd.DataFrame, meta: dict, output_path: str):
    """Writes the DataFrame to an Excel file with a summary sheet."""
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # --- Orders Sheet ---
        df.to_excel(writer, sheet_name='Order', index=False)
        order_ws = writer.sheets['Order']
        # Adjust column widths for the 'Order' sheet
        for col_idx, col in enumerate(df.columns, 1):
            column_letter = get_column_letter(col_idx)
            max_len = max(df[col].astype(str).map(len).max(), len(col))
            # Add some padding, especially for long descriptions
            adjusted_width = max_len + 4 if col == 'Description' else max_len + 2
            order_ws.column_dimensions[column_letter].width = adjusted_width

    logging.info(f"Excel file written to: {output_path}")
