import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from orders_converter.io.pdf_reader import read_pdf_table_and_meta
from orders_converter.io.excel_writer import write_order_excel
import pandas as pd

PRIMARY_COLOR = '#0066CC'

class OrdersSheetConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Orders Sheet Converter')
        self.geometry('520x320')
        self.resizable(False, False)
        self.configure(bg='white')
        self.pdf_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.filename = tk.StringVar()
        self.status = tk.StringVar(value='Ready')
        self._build_ui()

    def _build_ui(self):
        # Logo header (placeholder, replace with actual logo PNG)
        logo = tk.Label(self, text='[Logo]', bg='white', fg=PRIMARY_COLOR, font=('Arial', 20, 'bold'))
        logo.pack(pady=(16, 8))

        frm = ttk.Frame(self)
        frm.pack(padx=24, pady=4, fill='x')

        # File selector
        ttk.Label(frm, text='PDF File:').grid(row=0, column=0, sticky='w')
        file_entry = ttk.Entry(frm, textvariable=self.pdf_path, width=40)
        file_entry.grid(row=0, column=1, padx=4)
        file_btn = ttk.Button(frm, text='Browse...', command=self.browse_pdf)
        file_btn.grid(row=0, column=2)

        # Output location selector
        ttk.Label(frm, text='Output Folder:').grid(row=1, column=0, sticky='w')
        out_entry = ttk.Entry(frm, textvariable=self.output_dir, width=40)
        out_entry.grid(row=1, column=1, padx=4)
        out_btn = ttk.Button(frm, text='Browse...', command=self.browse_output_dir)
        out_btn.grid(row=1, column=2)

        # Filename field
        ttk.Label(frm, text='Excel Filename:').grid(row=2, column=0, sticky='w')
        fname_entry = ttk.Entry(frm, textvariable=self.filename, width=40)
        fname_entry.grid(row=2, column=1, padx=4)

        # Convert button
        self.convert_btn = ttk.Button(frm, text='Convert', command=self.convert, state='disabled')
        self.convert_btn.grid(row=3, column=1, pady=16)

        # Status bar
        status_bar = ttk.Label(self, textvariable=self.status, relief='sunken', anchor='w')
        status_bar.pack(side='bottom', fill='x')

        # Bind events
        self.pdf_path.trace_add('write', self._on_pdf_change)
        self.output_dir.trace_add('write', self._on_field_change)
        self.filename.trace_add('write', self._on_field_change)

    def browse_pdf(self):
        path = filedialog.askopenfilename(filetypes=[('PDF Files', '*.pdf')])
        if path:
            self.pdf_path.set(path)
            # Suggest output dir and filename
            self.output_dir.set(os.path.dirname(path))
            base = os.path.splitext(os.path.basename(path))[0]
            self.filename.set(base + '.xlsx')

    def browse_output_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir.set(path)

    def _on_pdf_change(self, *args):
        self._update_convert_btn()

    def _on_field_change(self, *args):
        self._update_convert_btn()

    def _update_convert_btn(self):
        if self.pdf_path.get() and self.output_dir.get() and self.filename.get():
            self.convert_btn.config(state='normal')
        else:
            self.convert_btn.config(state='disabled')

    def convert(self):
        pdf = self.pdf_path.get()
        out_dir = self.output_dir.get()
        fname = self.filename.get()
        out_path = os.path.join(out_dir, fname)
        try:
            meta, rows = read_pdf_table_and_meta(pdf)
            if not rows or len(rows) < 2:
                raise ValueError('No table rows found.')
            df = pd.DataFrame(rows[1:], columns=rows[0])
            write_order_excel(df, meta, out_path)
            self.status.set(f'Success: {out_path}')
        except Exception as e:
            self.status.set(f'Error: {e}')
            messagebox.showerror('Conversion Failed', str(e))

if __name__ == '__main__':
    app = OrdersSheetConverterApp()
    app.mainloop() 