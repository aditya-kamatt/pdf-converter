import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from orders_converter.io.pdf_reader import read_pdf_table_and_meta
from orders_converter.io.excel_writer import write_to_excel
import pandas as pd
import threading
import logging
from pathlib import Path
from orders_converter.utils.logging_config import setup_logging

# --- Setup Logging ---
setup_logging()
# ---

# Modern color scheme
PRIMARY_COLOR = '#2563EB'  # Blue
SECONDARY_COLOR = '#64748B'  # Slate
SUCCESS_COLOR = '#059669'  # Green
ERROR_COLOR = '#DC2626'  # Red
WARNING_COLOR = '#D97706'  # Amber
BG_COLOR = '#F8FAFC'  # Light gray background
CARD_BG = '#FFFFFF'  # White card background
TEXT_PRIMARY = '#1E293B'  # Dark text
TEXT_SECONDARY = '#64748B'  # Secondary text

class ModernButton(tk.Button):
    """Custom styled button with modern appearance"""
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.configure(
            relief='flat',
            borderwidth=0,
            font=('Segoe UI', 10, 'normal'),
            cursor='hand2'
        )
        self.bind('<Enter>', self._on_enter)
        self.bind('<Leave>', self._on_leave)
    
    def _on_enter(self, event):
        self.configure(relief='solid', borderwidth=1)
    
    def _on_leave(self, event):
        self.configure(relief='flat', borderwidth=0)

class FileDropFrame(tk.Frame):
    """Custom file drop area with visual feedback"""
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.configure(
            bg=CARD_BG,
            relief='solid',
            borderwidth=2,
            highlightthickness=0
        )
        
        self.label = tk.Label(
            self,
            text="📄 Drop PDF file here or click to browse",
            font=('Segoe UI', 12),
            fg=TEXT_SECONDARY,
            bg=CARD_BG
        )
        self.label.pack(expand=True, pady=20)
        
        self.bind('<Button-1>', self._on_click)
        self.label.bind('<Button-1>', self._on_click)
        
    def _on_click(self, event):
        self.event_generate('<<FileBrowse>>')

class OrdersSheetConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('PDF to Excel Converter')
        self.geometry('600x700')
        self.resizable(True, True)
        self.configure(bg=BG_COLOR)
        
        # Configure style
        self._configure_styles()
        
        # Variables
        self.pdf_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.filename = tk.StringVar()
        self.status = tk.StringVar(value='Ready to convert')
        self.progress_var = tk.DoubleVar()
        
        # Build UI
        self._build_ui()
        
        # Bind events
        self._bind_events()

    def _configure_styles(self):
        """Configure ttk styles for modern appearance"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure styles
        style.configure('Title.TLabel', 
                       font=('Segoe UI', 24, 'bold'), 
                       foreground=PRIMARY_COLOR,
                       background=BG_COLOR)
        
        style.configure('Subtitle.TLabel', 
                       font=('Segoe UI', 12), 
                       foreground=TEXT_SECONDARY,
                       background=BG_COLOR)
        
        style.configure('Card.TFrame', 
                       background=CARD_BG,
                       relief='solid',
                       borderwidth=1)
        
        style.configure('Success.TLabel',
                       font=('Segoe UI', 10),
                       foreground=SUCCESS_COLOR,
                       background=BG_COLOR)
        
        style.configure('Error.TLabel',
                       font=('Segoe UI', 10),
                       foreground=ERROR_COLOR,
                       background=BG_COLOR)

    def _build_ui(self):
        """Build the main UI components"""
        # Main container
        main_frame = tk.Frame(self, bg=BG_COLOR)
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Header
        self._build_header(main_frame)
        
        # File selection card
        self._build_file_selection(main_frame)
        
        # Output settings card
        self._build_output_settings(main_frame)
        
        # Convert button
        self._build_convert_section(main_frame)
        
        # Progress and status
        self._build_progress_status(main_frame)

    def _build_header(self, parent):
        """Build the header section"""
        header_frame = tk.Frame(parent, bg=BG_COLOR)
        header_frame.pack(fill='x', pady=(0, 20))
        
        # Title
        title_label = ttk.Label(
            header_frame,
            text="📊 PDF to Excel Converter",
            style='Title.TLabel'
        )
        title_label.pack()
        
        # Subtitle
        subtitle_label = ttk.Label(
            header_frame,
            text="Convert purchase order PDFs to structured Excel files",
            style='Subtitle.TLabel'
        )
        subtitle_label.pack(pady=(5, 0))

    def _build_file_selection(self, parent):
        """Build the file selection card"""
        card_frame = ttk.Frame(parent, style='Card.TFrame')
        card_frame.pack(fill='x', pady=(0, 15))
        
        # Card header
        card_header = tk.Frame(card_frame, bg=CARD_BG)
        card_header.pack(fill='x', padx=20, pady=(15, 10))
        
        header_label = tk.Label(
            card_header,
            text="📁 Select PDF File",
            font=('Segoe UI', 14, 'bold'),
            fg=TEXT_PRIMARY,
            bg=CARD_BG
        )
        header_label.pack(anchor='w')
        
        # File drop area
        self.file_drop = FileDropFrame(card_frame)
        self.file_drop.pack(fill='x', padx=20, pady=(0, 15))
        
        # Selected file display
        self.file_display = tk.Label(
            card_frame,
            textvariable=self.pdf_path,
            font=('Segoe UI', 10),
            fg=TEXT_PRIMARY,
            bg=CARD_BG,
            anchor='w',
            wraplength=500
        )
        self.file_display.pack(fill='x', padx=20, pady=(0, 15))

    def _build_output_settings(self, parent):
        """Build the output settings card"""
        card_frame = ttk.Frame(parent, style='Card.TFrame')
        card_frame.pack(fill='x', pady=(0, 15))
        
        # Card header
        card_header = tk.Frame(card_frame, bg=CARD_BG)
        card_header.pack(fill='x', padx=20, pady=(15, 10))
        
        header_label = tk.Label(
            card_header,
            text="💾 Output Settings",
            font=('Segoe UI', 14, 'bold'),
            fg=TEXT_PRIMARY,
            bg=CARD_BG
        )
        header_label.pack(anchor='w')
        
        # Settings frame
        settings_frame = tk.Frame(card_frame, bg=CARD_BG)
        settings_frame.pack(fill='x', padx=20, pady=(0, 15))
        
        # Output directory
        dir_frame = tk.Frame(settings_frame, bg=CARD_BG)
        dir_frame.pack(fill='x', pady=(0, 10))
        
        dir_label = tk.Label(
            dir_frame,
            text="Output Folder:",
            font=('Segoe UI', 10, 'bold'),
            fg=TEXT_PRIMARY,
            bg=CARD_BG
        )
        dir_label.pack(anchor='w')
        
        dir_entry_frame = tk.Frame(dir_frame, bg=CARD_BG)
        dir_entry_frame.pack(fill='x', pady=(5, 0))
        
        self.dir_entry = ttk.Entry(
            dir_entry_frame,
            textvariable=self.output_dir,
            font=('Segoe UI', 10)
        )
        self.dir_entry.pack(side='left', fill='x', expand=True)
        
        self.dir_btn = ModernButton(
            dir_entry_frame,
            text="Browse",
            bg=SECONDARY_COLOR,
            fg='white',
            command=self.browse_output_dir
        )
        self.dir_btn.pack(side='right', padx=(10, 0))
        
        # Filename
        filename_frame = tk.Frame(settings_frame, bg=CARD_BG)
        filename_frame.pack(fill='x', pady=(0, 10))
        
        filename_label = tk.Label(
            filename_frame,
            text="Excel Filename:",
            font=('Segoe UI', 10, 'bold'),
            fg=TEXT_PRIMARY,
            bg=CARD_BG
        )
        filename_label.pack(anchor='w')
        
        self.filename_entry = ttk.Entry(
            filename_frame,
            textvariable=self.filename,
            font=('Segoe UI', 10)
        )
        self.filename_entry.pack(fill='x', pady=(5, 0))

    def _build_convert_section(self, parent):
        """Build the convert button section"""
        convert_frame = tk.Frame(parent, bg=BG_COLOR)
        convert_frame.pack(fill='x', pady=(0, 15))
        
        self.convert_btn = ModernButton(
            convert_frame,
            text="🔄 Convert PDF to Excel",
            font=('Segoe UI', 14, 'bold'),
            bg=PRIMARY_COLOR,
            fg='white',
            command=self.convert,
            state='disabled',
            width=25,
            height=2
        )
        self.convert_btn.pack()

    def _build_progress_status(self, parent):
        """Build the progress and status section"""
        status_frame = tk.Frame(parent, bg=BG_COLOR)
        status_frame.pack(fill='x')
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(
            status_frame,
            variable=self.progress_var,
            mode='determinate',
            length=400
        )
        self.progress_bar.pack(pady=(0, 10))
        
        # Status label
        self.status_label = ttk.Label(
            status_frame,
            textvariable=self.status,
            style='Subtitle.TLabel'
        )
        self.status_label.pack()

    def _bind_events(self):
        """Bind UI events"""
        self.file_drop.bind('<<FileBrowse>>', self.browse_pdf)
        self.pdf_path.trace_add('write', self._on_pdf_change)
        self.output_dir.trace_add('write', self._on_field_change)
        self.filename.trace_add('write', self._on_field_change)

    def browse_pdf(self, event=None):
        """Browse for PDF file"""
        path = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[('PDF Files', '*.pdf'), ('All Files', '*.*')]
        )
        if path:
            self.pdf_path.set(path)
            # Suggest output dir and filename
            self.output_dir.set(os.path.dirname(path))
            base = os.path.splitext(os.path.basename(path))[0]
            self.filename.set(base + '.xlsx')
            
            # Update file display
            self.file_drop.label.configure(
                text=f"✅ {os.path.basename(path)}",
                fg=SUCCESS_COLOR
            )

    def browse_output_dir(self):
        """Browse for output directory"""
        path = filedialog.askdirectory(title="Select Output Folder")
        if path:
            self.output_dir.set(path)

    def _on_pdf_change(self, *args):
        """Handle PDF path change"""
        self._update_convert_btn()

    def _on_field_change(self, *args):
        """Handle field changes"""
        self._update_convert_btn()

    def _update_convert_btn(self):
        """Update convert button state"""
        if self.pdf_path.get() and self.output_dir.get() and self.filename.get():
            self.convert_btn.configure(state='normal', bg=PRIMARY_COLOR)
        else:
            self.convert_btn.configure(state='disabled', bg=SECONDARY_COLOR)

    def _update_progress(self, value, status):
        """Update progress bar and status"""
        self.progress_var.set(value)
        self.status.set(status)
        self.update_idletasks()

    def convert(self):
        """Convert PDF to Excel in a separate thread"""
        self.convert_btn.config(state='disabled', text="Converting...")
        self.status.set('Starting conversion...')
        self.status_label.config(style='TLabel')
        
        try:
            self._update_progress(10, 'Reading PDF...')
            
            pdf_path_str = self.pdf_path.get()
            output_dir_str = self.output_dir.get()
            filename_str = self.filename.get()
            
            logging.info(f"--- Conversion process started from GUI ---")
            logging.info(f"Input PDF: {pdf_path_str}")
            logging.info(f"Output Path: {os.path.join(output_dir_str, filename_str)}")
            
            meta_data, table_data = read_pdf_table_and_meta(pdf_path_str)
            
            self._update_progress(50, 'Extracting data from PDF...')
            
            if not table_data or len(table_data) <= 1:
                raise ValueError("No table data found in the PDF.")
            
            # Ensure we have the header and data properly separated
            if isinstance(table_data[0], list):
                header = table_data[0]
                data = table_data[1:]
            else:
                raise ValueError("Invalid table data format")
                
            df = pd.DataFrame(data, columns=header)
            
            logging.info(f"Created DataFrame with {len(df)} rows and {len(df.columns)} columns.")

            self._update_progress(80, 'Writing data to Excel...')
            
            output_path = os.path.join(output_dir_str, filename_str)
            write_to_excel(df, meta_data, output_path)
            
            self._update_progress(100, 'Conversion successful!')
            logging.info("--- Conversion process completed successfully ---")
            
            self.progress_var.set(100)
            self.status.set('✅ Conversion successful!')
            self.status_label.config(style='Success.TLabel')
            
            try:
                os.startfile(output_path)
            except AttributeError:
                logging.warning(f"Could not automatically open {output_path}. Please open it manually.")

        except Exception as e:
            self.progress_var.set(100)
            self.status.set(f"❌ Error: {e}")
            self.status_label.config(style='Error.TLabel')
            logging.error(f"An error occurred during conversion: {e}", exc_info=True)
            messagebox.showerror("Conversion Error", f"An error occurred: {e}")

        finally:
            self._update_progress(100, self.status.get())
            self.convert_btn['state'] = 'normal'
            self.convert_btn.config(text="Start Conversion")

if __name__ == '__main__':
    app = OrdersSheetConverterApp()
    app.mainloop() 