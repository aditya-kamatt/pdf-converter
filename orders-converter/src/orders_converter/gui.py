import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from orders_converter.io.pdf_reader import read_pdf_table_and_meta
from orders_converter.io.excel_writer import write_to_excel
import pandas as pd
import logging
from pathlib import Path
from orders_converter.utils.logging_config import setup_logging
from PIL import Image, ImageTk

# --- Setup Logging ---
setup_logging()
# ---

# --- Modern Theming ---
BG_COLOR = '#F8FAFC'  # Lighter gray
CARD_BG = '#FFFFFF'
PRIMARY_COLOR = '#2563EB'  # A modern blue
SECONDARY_COLOR = '#64748B'  # Slate gray
SUCCESS_COLOR = '#059669'  # Green
ERROR_COLOR = '#DC2626'  # Red
TEXT_PRIMARY = '#1E293B'  # Darker text
TEXT_SECONDARY = '#64748B'  # Lighter text
BORDER_COLOR = '#E2E8F0'

class ModernButton(tk.Button):
    def __init__(self, parent, **kwargs):
        self.hover_color = kwargs.pop('hover_color', '#1D4ED8') # Darker blue for hover
        self.original_bg = kwargs.get('bg', PRIMARY_COLOR)
        
        super().__init__(
            parent, 
            relief=tk.FLAT, 
            activebackground=self.hover_color, 
            activeforeground=kwargs.get('fg', 'white'),
            compound='left',
            **kwargs
        )
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _on_enter(self, event):
        if self['state'] == 'normal':
            self.config(bg=self.hover_color)

    def _on_leave(self, event):
        if self['state'] == 'normal':
            self.config(bg=self.original_bg)

class FileDropFrame(tk.Frame):
    """A frame that accepts file drops and looks good."""
    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=CARD_BG)
        self.config(
            highlightbackground=BORDER_COLOR, 
            highlightcolor=PRIMARY_COLOR, 
            highlightthickness=2, 
            bd=0
        )
        
        self.label = tk.Label(
            self, 
            text="📄 Drop PDF file here or click to browse", 
            font=('Segoe UI', 12), 
            bg=CARD_BG, 
            fg=TEXT_SECONDARY,
            pady=40
        )
        self.label.pack(fill='x', expand=True, padx=20)
        
        self.label.bind("<Button-1>", self._on_click)
        
    def _on_click(self, event):
        # Propagate the click to the parent to trigger the file dialog
        self.event_generate("<<FileBrowse>>")

class OrdersSheetConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Orders Sheet Converter")
        self.configure(bg=BG_COLOR)
        
        self.eval('tk::PlaceWindow . center')

        # --- Variables ---
        self.pdf_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.filename = tk.StringVar()
        self.status = tk.StringVar(value='')

        # --- Build UI ---
        self._build_ui()
        self._bind_events()

    def _build_ui(self):
        # Create a main frame that holds everything
        main_container = tk.Frame(self, bg=BG_COLOR)
        main_container.pack(fill='both', expand=True)

        # --- Footer for Button and Status ---
        # This frame is packed first to the bottom and does not expand
        footer_frame = tk.Frame(main_container, bg=BG_COLOR, padx=30, pady=20)
        footer_frame.pack(side='bottom', fill='x', expand=False)
        self._build_convert_section(footer_frame)
        self._build_progress_status(footer_frame)

        # --- Main Content Area ---
        # This frame holds the rest of the content and expands
        content_frame = tk.Frame(main_container, bg=BG_COLOR, padx=30, pady=30)
        content_frame.pack(side='top', fill='both', expand=True)
        
        self._build_header(content_frame)
        self._build_file_selection(content_frame)
        self._build_output_settings(content_frame)

    def _build_header(self, parent):
        header_frame = tk.Frame(parent, bg=BG_COLOR)
        header_frame.pack(fill='x', pady=(0, 30))
        
        title_frame = tk.Frame(header_frame, bg=BG_COLOR)
        title_frame.pack(side='left')

        tk.Label(
            title_frame,
            text="Orders Sheet Converter",
            font=('Segoe UI', 24, 'bold'),
            fg=TEXT_PRIMARY,
            bg=BG_COLOR
        ).pack(anchor='w')

        # --- Logo ---
        try:
            base_path = Path(__file__).resolve().parent.parent.parent
            logo_path = base_path / "assets" / "logo.jpg"
            if not logo_path.exists():
                logo_path = base_path / "assets" / "logo.png"

            if logo_path.exists():
                pil_image = Image.open(logo_path).resize((50, 50), Image.Resampling.LANCZOS)
                self.logo_image = ImageTk.PhotoImage(pil_image)
                logo_label = tk.Label(header_frame, image=self.logo_image, bg=BG_COLOR)
                logo_label.pack(side='right', padx=(20, 0))
            else:
                logging.warning("Logo not found in assets folder.")
        except Exception as e:
            logging.error(f"Failed to load logo: {e}")

    def _build_file_selection(self, parent):
        frame = tk.Frame(parent, bg=CARD_BG, relief='solid', bd=1, highlightbackground=BORDER_COLOR)
        frame.pack(fill='x', expand=False, pady=(0, 20))
        
        header_frame = tk.Frame(frame, bg=CARD_BG)
        header_frame.pack(fill='x', padx=20, pady=15)
        
        tk.Label(
            header_frame,
            text="Select PDF File",
            font=('Segoe UI', 14, 'bold'),
            fg=TEXT_PRIMARY,
            bg=CARD_BG
        ).pack(side='left')

        self.file_drop = FileDropFrame(frame)
        self.file_drop.pack(fill='x', expand=False, padx=20, pady=(0, 20))

    def _build_output_settings(self, parent):
        frame = tk.Frame(parent, bg=CARD_BG, relief='solid', bd=1, highlightbackground=BORDER_COLOR)
        frame.pack(fill='x', expand=False, pady=0)

        header_frame = tk.Frame(frame, bg=CARD_BG)
        header_frame.pack(fill='x', padx=20, pady=15)
        
        tk.Label(
            header_frame,
            text="Output Settings",
            font=('Segoe UI', 14, 'bold'),
            fg=TEXT_PRIMARY,
            bg=CARD_BG
        ).pack(side='left')
        
        settings_frame = tk.Frame(frame, bg=CARD_BG)
        settings_frame.pack(fill='x', expand=False, padx=20, pady=20)

        # Output Folder
        dir_frame = tk.Frame(settings_frame, bg=CARD_BG)
        dir_frame.pack(fill='x', pady=(0, 15))
        
        tk.Label(
            dir_frame,
            text="Output Folder:",
            font=('Segoe UI', 10, 'bold'),
            fg=TEXT_PRIMARY,
            bg=CARD_BG
        ).pack(anchor='w')
        
        dir_entry_frame = tk.Frame(dir_frame, bg=CARD_BG)
        dir_entry_frame.pack(fill='x', pady=(5, 0))
        
        self.dir_entry = ttk.Entry(
            dir_entry_frame,
            textvariable=self.output_dir,
            font=('Segoe UI', 10)
        )
        self.dir_entry.pack(side='left', fill='x', expand=True, ipady=4)
        
        self.dir_btn = ModernButton(
            dir_entry_frame,
            text="Browse",
            bg=SECONDARY_COLOR,
            hover_color='#5A6268',
            fg='white',
            font=('Segoe UI', 9, 'bold'),
            command=self.browse_output_dir
        )
        self.dir_btn.pack(side='right', padx=(10, 0))
        
        # Filename
        filename_frame = tk.Frame(settings_frame, bg=CARD_BG)
        filename_frame.pack(fill='x')
        
        tk.Label(
            filename_frame,
            text="Excel Filename:",
            font=('Segoe UI', 10, 'bold'),
            fg=TEXT_PRIMARY,
            bg=CARD_BG
        ).pack(anchor='w')
        
        self.filename_entry = ttk.Entry(
            filename_frame,
            textvariable=self.filename,
            font=('Segoe UI', 10)
        )
        self.filename_entry.pack(fill='x', pady=(5, 0), ipady=4)

    def _build_convert_section(self, parent):
        # This function now packs into the footer_frame
        self.convert_btn = ModernButton(
            parent, # Pack directly into the parent (footer)
            text="🔄 Convert to Excel",
            font=('Segoe UI', 14, 'bold'),
            bg=PRIMARY_COLOR,
            fg='#FFFFFF',
            disabledforeground='#B0BEC5',  # Light blue-gray for disabled text
            command=self.convert,
            state='disabled',
            height=2,
            width=25,
            bd=0
        )
        self.convert_btn.pack()

    def _build_progress_status(self, parent):
        # This function now packs into the footer_frame
        self.status_label = ttk.Label(
            parent, # Pack directly into the parent (footer)
            textvariable=self.status,
            font=('Segoe UI', 10),
            foreground=TEXT_SECONDARY,
            background=BG_COLOR
        )
        self.status_label.pack(pady=(10, 0))

    def _bind_events(self):
        self.file_drop.bind('<<FileBrowse>>', self.browse_pdf)
        self.pdf_path.trace_add('write', self._on_field_change)
        self.output_dir.trace_add('write', self._on_field_change)
        self.filename.trace_add('write', self._on_field_change)

    def browse_pdf(self, event=None):
        path = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[('PDF Files', '*.pdf')]
        )
        if path:
            self.pdf_path.set(path)
            self.output_dir.set(os.path.dirname(path))
            base = os.path.splitext(os.path.basename(path))[0]
            self.filename.set(base + '.xlsx')
            
            self.file_drop.label.configure(
                text=f"✅ {os.path.basename(path)}",
                fg=SUCCESS_COLOR
            )

    def browse_output_dir(self):
        path = filedialog.askdirectory(title="Select Output Folder")
        if path:
            self.output_dir.set(path)

    def _on_field_change(self, *args):
        if self.pdf_path.get() and self.output_dir.get() and self.filename.get():
            self.convert_btn.config(state='normal')
        else:
            self.convert_btn.config(state='disabled')

    def _update_status(self, text, color=TEXT_SECONDARY):
        self.status.set(text)
        self.status_label.config(foreground=color)
        self.update_idletasks()

    def convert(self):
        self.convert_btn.config(state='disabled', text="Converting...")
        self._update_status('Starting conversion...')
        
        try:
            pdf_path_str = self.pdf_path.get()
            output_dir_str = self.output_dir.get()
            filename_str = self.filename.get()
            
            logging.info("--- Conversion process started from GUI ---")
            logging.info(f"Input PDF: {pdf_path_str}")
            
            meta_data, table_data = read_pdf_table_and_meta(pdf_path_str)
            
            if not table_data or len(table_data) <= 1:
                raise ValueError("No table data found in the PDF.")
            
            header, data = table_data[0], table_data[1:]
            df = pd.DataFrame(data, columns=header)
            logging.info(f"Created DataFrame with {len(df)} rows.")

            output_path = os.path.join(output_dir_str, filename_str)
            write_to_excel(df, meta_data, output_path)
            
            self._update_status('✅ Conversion successful!', SUCCESS_COLOR)
            
            try:
                os.startfile(output_path)
            except AttributeError:
                logging.warning(f"Could not open {output_path}. Please open it manually.")

        except Exception as e:
            self._update_status(f"❌ Error: {e}", ERROR_COLOR)
            logging.error(f"Conversion error: {e}", exc_info=True)
            messagebox.showerror("Conversion Error", f"An error occurred: {e}")

        finally:
            self.convert_btn.config(text="🔄 Convert to Excel")
            self._on_field_change()

def main():
    """Main function to run the application."""
    app = OrdersSheetConverterApp()
    app.mainloop()

if __name__ == '__main__':
    main() 