import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
def show_splash(root, splash_path, min_time=2000):
    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    splash.lift()
    splash.attributes("-topmost", True)
    # Load splash image
    img = Image.open(splash_path)
    img = img.convert("RGBA")
    splash_img = ImageTk.PhotoImage(img)
    w, h = splash_img.width(), splash_img.height()
    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    x = (screen_w - w) // 2
    y = (screen_h - h) // 3
    splash.geometry(f"{w}x{h}+{x}+{y}")
    label = tk.Label(splash, image=splash_img, borderwidth=0)
    label.image = splash_img
    label.pack()
    root.withdraw()
    def close_splash():
        splash.destroy()
        root.deiconify()
    root.after(min_time, close_splash)
    root.update()
from pathlib import Path
import sys
import ctypes
import PyPDF2
from typing import List, Dict, Optional
from datetime import datetime
import os
import webbrowser
import json
from PIL import Image, ImageTk
import io
import fitz  # PyMuPDF
import threading

__VERSION__ = "1.5.0"


def _enable_dpi_awareness() -> None:
    if os.name != 'nt':
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

class ToolTip:
    """Create a tooltip for a given widget"""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)
    
    def show_tooltip(self, event=None):
        if self.tooltip_window or not self.text:
            return
        x, y, _, _ = self.widget.bbox("insert") if hasattr(self.widget, 'bbox') else (0, 0, 0, 0)
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        
        self.tooltip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        
        label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                        background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                        font=("Arial", 9), padx=8, pady=6)
        label.pack()
    
    def hide_tooltip(self, event=None):
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None

class PDFCombinerApp:
    def _get_dpi_scale(self) -> float:
        try:
            return self.root.winfo_fpixels("1i") / 96.0
        except Exception:
            return 1.0

    def _scale_geometry(self, width: int, height: int) -> tuple[int, int]:
        scale = self._get_dpi_scale()
        return max(1, int(width * scale)), max(1, int(height * scale))

    def __init__(self, root):
        self.root = root
        self.root.title("PDF Combiner")
        
        # Center window horizontally and align to top
        window_width, window_height = self._scale_geometry(740, 540)
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        center_x = int((screen_width - window_width) / 2)
        center_y = 5
        self.root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        self.root.resizable(False, False)
        
        # Helper to get resource path both when running normally and when frozen
        def resource_path(relative_path: str) -> str:
            if getattr(sys, "frozen", False):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(__file__)
            return os.path.join(base_path, relative_path)

        # Use existing PNG icon if available, otherwise fallback to .ico
        try:
            png_path = resource_path("pdfcombinericon.png")
            icon_image = tk.PhotoImage(file=png_path)
            self.root.iconphoto(True, icon_image)
            # keep a reference so it isn't garbage-collected
            self._icon_image = icon_image
        except Exception:
            try:
                ico_path = resource_path("pdfcombinericon.ico")
                self.root.iconbitmap(ico_path)
            except Exception:
                pass
        
        # Data structure: list of dicts with 'path', 'rotation', and 'page_range' keys
        self.pdf_files: List[Dict[str, any]] = []
        self.combine_order = tk.StringVar(value="display")
        self.drag_start_index = None
        self.drag_start_y = None
        self.is_dragging = False
        self.auto_scroll_id = None  # Track auto-scroll timer during drag
        self.updating_visuals = False  # Flag to prevent configure during visual updates
        self.last_scrollregion = None  # Track last scrollregion to avoid unnecessary updates
        self.row_visual_state = {}  # Track last background color of each row to avoid unnecessary updates
        # Prefer Documents when it exists; otherwise fall back to a known existing folder.
        home_dir = Path.home()
        candidates = [home_dir / "Documents", home_dir / "Desktop", home_dir]
        default_dir = next((p for p in candidates if p.exists()), home_dir)
        self.output_directory = str(default_dir)
        self.add_files_directory = str(default_dir)  # Default directory for adding files
        self.list_files_directory = str(default_dir)  # Default directory for load/save list
        self.output_filename = tk.StringVar(value="combined.pdf")
        self.last_output_file = None
        self.preview_window = None
        self.preview_file_index = None
        self.preview_label = None
        self.preview_after_id = None
        self.preview_delay_ms = 400
        self.pending_preview_index = None
        self.preview_enabled = tk.BooleanVar(value=True)  # Preview on hover enabled by default
        self.add_filename_bookmarks = tk.BooleanVar(value=True)  # Add filename bookmarks enabled by default
        self.insert_blank_pages = tk.BooleanVar(value=False)  # Insert breaker pages between files
        self.breaker_pages_uniform_size = tk.BooleanVar(value=False)  # Make breaker pages uniform size
        self.insert_toc = tk.BooleanVar(value=True)  # Insert table of contents page
        self.rotation_vars = {}  # Map of index to tk.StringVar for rotation dropdowns
        self.page_range_vars = {}  # Map of index to tk.StringVar for page ranges
        self.page_range_last_valid = {}  # Track last valid page range per index
        self.reverse_vars = {}  # Map of index to tk.BooleanVar for page reversal
        
        # New advanced features
        self.compression_quality = tk.StringVar(value="None")  # Compression level
        self.enable_metadata = tk.BooleanVar(value=False)  # Enable metadata editing
        self.pdf_title = tk.StringVar(value="")  # Metadata: title
        self.pdf_author = tk.StringVar(value="")  # Metadata: author
        self.pdf_subject = tk.StringVar(value="")  # Metadata: subject 
        self.pdf_keywords = tk.StringVar(value="")  # Metadata: keywords
        self.enable_page_scaling = tk.BooleanVar(value=False)  # Scale to uniform size
        self.enable_watermark = tk.BooleanVar(value=False)  # Add watermark
        self.watermark_text = tk.StringVar(value="")  # Watermark text
        self.watermark_opacity = tk.DoubleVar(value=0.3)  # Watermark opacity (0.1-0.9)
        self.watermark_font_size = tk.IntVar(value=50)  # Watermark font size
        self.watermark_rotation = tk.IntVar(value=45)  # Watermark rotation (0-360 degrees)
        self.watermark_position = tk.StringVar(value="Center")  # Watermark position (top, center, bottom)
        self.watermark_safe_mode = tk.BooleanVar(value=True)  # Safe Mode: auto-adjust watermark to prevent clipping
        self.delete_blank_pages = tk.BooleanVar(value=False)  # Remove blank pages
        
        # Store last used metadata values
        self.last_metadata = {
            'title': '',
            'author': '',
            'subject': '',
            'keywords': ''
        }
        
        # Set config file location to AppData\Roaming\PDFCombiner on Windows
        if os.name == 'nt' and 'APPDATA' in os.environ:
            config_dir = Path(os.environ['APPDATA']) / "PDFCombiner"
        else:
            # Fallback for other platforms
            config_dir = Path.home() / ".pdfcombiner"
        self.config_file = config_dir / "config.json"
        
        # Load saved settings
        self._load_settings()
        
        # Configure custom style for notebook tabs
        style = ttk.Style()
        style.theme_use('clam')  # Use clam theme as base for better customization
        
        # Configure the notebook and tab appearance
        style.configure('TNotebook', background='#E0E0E0', borderwidth=2, relief='solid')
        style.configure('TNotebook.Tab', padding=[10, 4], font=('Arial', 10, 'bold'), background='#D0D0D0', foreground='#333333', focuscolor='#D0D0D0', width=16, anchor='center')
        style.map('TNotebook.Tab', 
                  background=[('selected', '#4A90E2'), ('active', '#5B9FE8')],
                  foreground=[('selected', 'white'), ('active', 'white')],
                  padding=[('selected', [10, 4])])
        
        # Configure Combobox style to match file list background
        style.configure('TCombobox', 
                       fieldbackground='white', 
                       background='white', 
                       foreground='black', 
                       selectbackground='white',
                       selectforeground='black',
                       relief='flat', 
                       borderwidth=0)
        style.map('TCombobox',
                  fieldbackground=[('readonly', 'white'), ('disabled', 'white'), ('focus', 'white'), ('!focus', 'white')],
                  background=[('readonly', 'white'), ('disabled', 'white')],
                  selectbackground=[('readonly', 'white'), ('disabled', 'white'), ('focus', 'white'), ('!focus', 'white')],
                  selectforeground=[('readonly', 'black'), ('disabled', 'gray'), ('focus', 'black'), ('!focus', 'black')])
        
        # Create notebook (tabbed interface)
        self.notebook = ttk.Notebook(root, style='TNotebook')
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # ===== INPUT TAB =====
        input_frame = tk.Frame(self.notebook)
        self.notebook.add(input_frame, text="Input Files")
        
        # Spacer to move content down
        spacer_frame = tk.Frame(input_frame, height=10)
        spacer_frame.pack()
        spacer_frame.pack_propagate(False)
        
        # Main title with shadow effect using Canvas for better control
        title_container = tk.Canvas(input_frame, width=280, height=30, bg=input_frame.cget('bg'), highlightthickness=0)
        title_container.pack(pady=(0, 0))
        
        # Draw shadow text (offset, subtle light gray)
        title_container.create_text(142, 17, text="PDF Combiner", font=("Arial", 14, "bold"), 
                                   fill="#BBBBBB", anchor="center")
        
        # Draw main title text (blue) 
        title_container.create_text(141, 16, text="PDF Combiner", font=("Arial", 14, "bold"), 
                                   fill="#000000", anchor="center")
        
        # Title and preview checkbox frame - same line
        title_frame = tk.Frame(input_frame)
        title_frame.pack(anchor=tk.W, fill=tk.X, padx=10, pady=(2, 5))
        
        # Title above list
        title_label = tk.Label(title_frame, text="List and Order of PDFs/Images to Combine:", font=("Arial", 10, "bold"))
        title_label.pack(side=tk.LEFT, anchor=tk.W)
        
        # Preview on hover checkbox
        preview_checkbox = tk.Checkbutton(
            title_frame,
            text="Preview first page on hover",
            variable=self.preview_enabled,
            command=self._on_preview_toggle,
            font=("Arial", 9)
        )
        preview_checkbox.pack(side=tk.RIGHT, padx=(10, 0))
        
        # Custom scrollable frame for file list with rotation controls
        list_frame = tk.Frame(input_frame)
        list_frame.pack(pady=(0, 10), padx=10, fill=tk.X)
        
        # Column headers using fixed-width labels
        header_frame = tk.Frame(list_frame, bg="#E0E0E0")
        header_frame.pack(anchor=tk.W, fill=tk.X)

        hdr_font = ("Consolas", 8)
        # Numbering column header
        num_hdr = tk.Label(header_frame, text="#", font=hdr_font, bg="#E0E0E0", width=4, anchor='e')
        num_hdr.pack(side=tk.LEFT, padx=(0, 2))

        # Filename header - clickable
        self.filename_hdr = tk.Label(header_frame, text="Filename", font=hdr_font, bg="#E0E0E0", width=72, anchor='w')
        self.filename_hdr.pack(side=tk.LEFT)
        self.filename_hdr.bind("<Button-1>", lambda e: self.on_sort_clicked('name'))
        self.filename_hdr.bind("<Enter>", lambda e: self.filename_hdr.config(cursor="hand2"))
        self.filename_hdr.bind("<Leave>", lambda e: self.filename_hdr.config(cursor="arrow"))

        # File Size header - clickable
        self.size_hdr = tk.Label(header_frame, text="Size", font=hdr_font, bg="#E0E0E0", width=11, anchor='w')
        self.size_hdr.pack(side=tk.LEFT)
        self.size_hdr.bind("<Button-1>", lambda e: self.on_sort_clicked('size'))
        self.size_hdr.bind("<Enter>", lambda e: self.size_hdr.config(cursor="hand2"))
        self.size_hdr.bind("<Leave>", lambda e: self.size_hdr.config(cursor="arrow"))

        # Date header - clickable
        self.date_hdr = tk.Label(header_frame, text="Date", font=hdr_font, bg="#E0E0E0", width=12, anchor='w')
        self.date_hdr.pack(side=tk.LEFT)
        self.date_hdr.bind("<Button-1>", lambda e: self.on_sort_clicked('date'))
        self.date_hdr.bind("<Enter>", lambda e: self.date_hdr.config(cursor="hand2"))
        self.date_hdr.bind("<Leave>", lambda e: self.date_hdr.config(cursor="arrow"))
        
        pages_hdr = tk.Label(header_frame, text="Pages", font=hdr_font, bg="#E0E0E0", width=7, anchor='w')
        pages_hdr.pack(side=tk.LEFT, padx=(4, 0))
        ToolTip(pages_hdr, "Specify page range to include from this PDF.\nExamples: '1-5', '1,3,5', '1-3,7-9'\nLeave blank or type All to include all pages.")
        
        rot_hdr = tk.Label(header_frame, text="Rot", font=hdr_font, bg="#E0E0E0", width=7, anchor='c')
        rot_hdr.pack(side=tk.LEFT, padx=2)
        ToolTip(rot_hdr, "Rotate all pages in this PDF.\nOptions: 0°, 90°, 180°, 270°\nclockwise")
        
        rev_hdr = tk.Label(header_frame, text="Rev", font=hdr_font, bg="#E0E0E0", width=4, anchor='c')
        rev_hdr.pack(side=tk.LEFT, padx=2)
        ToolTip(rev_hdr, "Reverse the page order of this PDF.\nLast page becomes first, first becomes last.")
        
        # Sub-frame for custom list frame and scrollbar (sized for ~11 rows)
        listbox_scroll_frame = tk.Frame(
            list_frame, 
            height=270,
            bd=0,
            relief=tk.FLAT,
            bg="white",
            highlightbackground="#CCCCCC",
            highlightthickness=1
        )
        listbox_scroll_frame.pack(fill=tk.X)
        listbox_scroll_frame.pack_propagate(False)  # Prevent children from resizing frame
        
        # Scrollbar
        scrollbar = tk.Scrollbar(listbox_scroll_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Canvas for scrolling - disable all focus highlighting
        self.file_list_canvas = tk.Canvas(listbox_scroll_frame, yscrollcommand=scrollbar.set, bg="white", 
                                          highlightthickness=0, bd=0, takefocus=0)
        self.file_list_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_list_canvas.yview)
        
        # Keep focus on canvas to prevent focus-change flicker
        self.file_list_canvas.focus_set()
        
        # Inner frame for content
        self.file_list_frame = tk.Frame(self.file_list_canvas, bg="white")
        canvas_window = self.file_list_canvas.create_window((0, 0), window=self.file_list_frame, anchor="nw")
        
        # Store last scrollregion to avoid unnecessary updates
        self.last_scrollregion = None
        
        # Configure scrollbar region - only when size actually changes
        def on_frame_configure(event=None):
            if self.updating_visuals:
                return  # Skip during visual-only updates
            
            # Get canvas dimensions
            canvas_width = self.file_list_canvas.winfo_width()
            
            # Update frame width to match canvas
            if canvas_width > 1:
                self.file_list_frame.configure(width=canvas_width)
                self.file_list_canvas.itemconfig(canvas_window, width=canvas_width)
            
            # Update the scrollregion to encompass the frame
            self.file_list_frame.update_idletasks()
            frame_width = self.file_list_frame.winfo_reqwidth()
            frame_height = self.file_list_frame.winfo_reqheight()
            
            canvas_height = self.file_list_canvas.winfo_height()
            if canvas_height <= 1:
                canvas_height = 270  # Default height fallback
            
            # Ensure scrollregion is at least as tall as canvas to prevent blank lines when scrolling
            scrollregion_height = max(frame_height, canvas_height)
            new_region = (0, 0, max(frame_width, 1), scrollregion_height)
            
            if new_region != self.last_scrollregion:
                self.last_scrollregion = new_region
                self.file_list_canvas.configure(scrollregion=new_region)
        
        # Store configure function for manual calls
        self.canvas_configure = on_frame_configure
        
        # Only bind to canvas resize, not frame configure (to avoid color change triggers)
        self.file_list_canvas.bind("<Configure>", lambda e: on_frame_configure())
        
        # Bind mousewheel for scrolling
        def _on_mousewheel(event):
            self.file_list_canvas.yview_scroll(-1 * (event.delta // 120), "units")
        
        self.file_list_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Store references
        self.canvas = self.file_list_canvas
        self.scrollbar = scrollbar
        self.file_listbox = None  # No legacy listbox anymore
        
        # File count label with legend
        count_legend_frame = tk.Frame(input_frame)
        count_legend_frame.pack(pady=1, fill=tk.X)
        
        # Legend on the far left
        legend_frame = tk.Frame(count_legend_frame)
        legend_frame.pack(side=tk.LEFT, padx=(10, 0))
        
        # Blue square for Image
        image_square = tk.Canvas(legend_frame, width=12, height=12, bg="SystemButtonFace", highlightthickness=0)
        image_square.pack(side=tk.LEFT, padx=(0, 2))
        image_square.create_rectangle(1, 1, 11, 11, fill="#4A90E2", outline="#4A90E2")
        
        image_label = tk.Label(legend_frame, text="Image", font=("Arial", 8))
        image_label.pack(side=tk.LEFT, padx=(0, 10))
        
        # Black square for PDF
        pdf_square = tk.Canvas(legend_frame, width=12, height=12, bg="SystemButtonFace", highlightthickness=0)
        pdf_square.pack(side=tk.LEFT, padx=(0, 2))
        pdf_square.create_rectangle(1, 1, 11, 11, fill="black", outline="black")
        
        pdf_label = tk.Label(legend_frame, text="PDF", font=("Arial", 8))
        pdf_label.pack(side=tk.LEFT)
        
        # File count centered
        self.count_label = tk.Label(count_legend_frame, text="Files to combine: 0", font=("Arial", 9))
        self.count_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        
        # Drag and drop instruction
        drag_drop_note = tk.Label(
            input_frame,
            text="After adding files, single click to select a file. Ctrl-Click to select multiple files. Click and drag files to reorder.\nHover to preview PDFs or images. Double-click to open a file. Click filename, size, or date column headers to sort.",
            font=("Arial", 8),
            fg="#666666"
        )
        drag_drop_note.pack(pady=1)
        
        # Sorting state
        self.sort_key = None  # 'name' | 'size' | 'date'
        self.sort_reverse = False
        
        # Button frame below listbox for file management buttons
        listbox_button_frame = tk.Frame(input_frame)
        listbox_button_frame.pack(pady=8)
        
        # Add files button
        self.add_button = tk.Button(
            listbox_button_frame,
            text="Add PDFs/Images...",
            command=self.add_files,
            width=18,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10)
        )
        self.add_button.grid(row=0, column=0, padx=5)
        
        # Remove selected button
        self.remove_button = tk.Button(
            listbox_button_frame,
            text="Remove Selected",
            command=self.remove_file,
            width=18,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10),
            state=tk.DISABLED  # Start disabled since list is empty
        )
        self.remove_button.grid(row=0, column=1, padx=5)
        
        # Clear all button
        self.clear_button = tk.Button(
            listbox_button_frame,
            text="Remove All",
            command=self.clear_files,
            width=18,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10),
            state=tk.DISABLED  # Start disabled since list is empty
        )
        self.clear_button.grid(row=0, column=2, padx=5)
        
        # Load/Save List button
        self.load_save_button = tk.Button(
            listbox_button_frame,
            text="Load/Save List..",
            command=self.show_load_save_dialog,
            width=18,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10)
        )
        self.load_save_button.grid(row=0, column=3, padx=5)
        
        # ===== OUTPUT TAB =====
        output_frame_main = tk.Frame(self.notebook)
        self.notebook.add(output_frame_main, text="Output Settings")
        
        # Padding frame for better spacing
        output_content_frame = tk.Frame(output_frame_main)
        output_content_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(10, 12))
        
        # Output settings frame
        output_frame = tk.LabelFrame(output_content_frame, text="Output Settings", font=("Arial", 10, "bold"), padx=10, pady=5)
        output_frame.pack(pady=(0, 3), fill=tk.X)
        
        # Filename frame
        filename_frame = tk.Frame(output_frame)
        filename_frame.pack(fill=tk.X, pady=2)
        
        tk.Label(filename_frame, text="Filename for combined PDF:", font=("Arial", 9)).pack(side=tk.LEFT, padx=5)
        
        # Create a bordered frame around the filename entry
        filename_box = tk.Frame(
            filename_frame,
            bd=0,
            relief=tk.FLAT,
            bg="white",
            highlightbackground="#CCCCCC",
            highlightthickness=1,
            padx=2,
            pady=1,
        )
        filename_box.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        filename_entry = tk.Entry(filename_box, textvariable=self.output_filename, font=("Arial", 9), width=30, border=0, bg="white")
        filename_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        filename_entry.bind("<FocusOut>", lambda e: self._validate_filename_on_focus_out())
        
        #tk.Label(filename_frame, text=".pdf", font=("Arial", 9)).pack(side=tk.LEFT)
        
        # Location frame (boxed to highlight save location)
        location_frame = tk.Frame(output_frame)
        location_frame.pack(fill=tk.X, pady=3)
        
        tk.Label(location_frame, text="Save Location:", font=("Arial", 9)).pack(side=tk.LEFT, padx=5)
        
        browse_button = tk.Button(
            location_frame,
            text="Browse",
            command=self.browse_output_location,
            width=10,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 9)
        )
        # Small bordered box around the save-location text only (placed left)
        # Use a thin highlight border for a slimmer look
        dir_box = tk.Frame(
            location_frame,
            bd=0,
            relief=tk.FLAT,
            bg="#FAFAFA",
            highlightbackground="#BBBBBB",
            highlightthickness=1,
            padx=2,
            pady=1,
        )
        dir_box.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        self.location_label = tk.Label(
            dir_box,
            text=self.output_directory,
            font=("Arial", 9),
            fg="#000",
            bg="#FAFAFA",
            anchor="w"
        )
        self.location_label.pack(side=tk.LEFT)
        # Pack the Browse button to the right of the save-location box
        browse_button.pack(side=tk.LEFT, padx=5)
        
        # Options frame
        options_frame = tk.LabelFrame(output_content_frame, text="Options", font=("Arial", 9, "bold"))
        options_frame.pack(pady=(0, 0), padx=0, fill=tk.X)

        # Bookmark and breaker page checkboxes on same row
        checkbox_row = tk.Frame(options_frame)
        checkbox_row.pack(fill=tk.X, pady=(0, 2), padx=8)
        checkbox_row.columnconfigure(0, weight=0)
        checkbox_row.columnconfigure(1, weight=1)
        
        bookmark_frame = tk.Frame(checkbox_row)
        bookmark_frame.grid(row=0, column=0, sticky="nw")
        
        bookmark_checkbox = tk.Checkbutton(
            bookmark_frame,
            text="Add filename bookmarks to the combined PDF",
            variable=self.add_filename_bookmarks,
            command=self._save_settings,
            font=("Arial", 9)
        )
        bookmark_checkbox.pack(anchor="w")
        ToolTip(bookmark_checkbox, "Adds each file's name as a bookmark in the combined\nPDF, linking to the start of that file's content.\nNote: Source PDF bookmarks are not currently preserved.")

        scale_checkbox = tk.Checkbutton(
            bookmark_frame,
            text="Scale all pages to uniform size",
            variable=self.enable_page_scaling,
            command=self._save_settings,
            font=("Arial", 9)
        )
        scale_checkbox.pack(anchor="w", pady=(2, 0))
        ToolTip(scale_checkbox, "Pages will be resized to a uniform size.")

        toc_checkbox = tk.Checkbutton(
            bookmark_frame,
            text="Insert Table of Contents page",
            variable=self.insert_toc,
            command=self._save_settings,
            font=("Arial", 9)
        )
        toc_checkbox.pack(anchor="w", pady=(2, 0))
        ToolTip(toc_checkbox, "Adds a clickable TOC page at the beginning\nwith links to each merged file")

        compression_frame = tk.Frame(bookmark_frame)
        compression_frame.pack(anchor="w", pady=(2, 0))
        tk.Label(compression_frame, text="Compression Level:", font=("Arial", 9)).pack(side=tk.LEFT, padx=(0, 5))
        compression_combo = ttk.Combobox(
            compression_frame,
            textvariable=self.compression_quality,
            values=["None", "Low", "Medium", "High", "Maximum"],
            width=12,
            state="readonly",
            font=("Arial", 9)
        )
        compression_combo.pack(side=tk.LEFT)
        compression_combo.bind("<<ComboboxSelected>>", lambda e: self._save_settings())
        compression_combo.bind("<FocusOut>", lambda e: self._validate_compression_quality())
        ToolTip(compression_combo, "Compresses the combined PDF.\nHigher compression = Smaller filesize but lowers quality")

        breaker_options_frame = tk.Frame(
            checkbox_row,
            bg=checkbox_row.cget("bg"),
            padx=6,
            pady=4
        )
        breaker_options_frame.grid(row=0, column=1, sticky="nw", padx=(15, 0))

        blank_pages_checkbox = tk.Checkbutton(
            breaker_options_frame,
            text="Insert breaker pages between files",
            variable=self.insert_blank_pages,
            command=self._toggle_breaker_page_options,
            font=("Arial", 9)
        )
        blank_pages_checkbox.pack(anchor="w")
        ToolTip(blank_pages_checkbox, "Insert a page before each combined file\nwith the filename shown")
        
        # Suboption for breaker pages: uniform size
        self.breaker_uniform_checkbox = tk.Checkbutton(
            breaker_options_frame,
            text="Make breaker pages a consistent size",
            variable=self.breaker_pages_uniform_size,
            command=self._save_settings,
            font=("Arial", 9),
            state="disabled"
        )
        self.breaker_uniform_checkbox.pack(anchor="w", padx=(18, 0))
        ToolTip(self.breaker_uniform_checkbox, "When enabled, all breaker pages will be standard letter size.\nWhen disabled, breaker pages match the following content size.")
        self._toggle_breaker_page_options()

        blank_detect_checkbox = tk.Checkbutton(
            breaker_options_frame,
            text="Ignore blank pages in source files when combining",
            variable=self.delete_blank_pages,
            command=self._save_settings,
            font=("Arial", 9)
        )
        blank_detect_checkbox.pack(anchor="w", pady=(4, 0))
        ToolTip(blank_detect_checkbox, "When combining all pages: Blank pages will be skipped.\nWhen selecting specific page range(s): All pages in the range\nwill be kept even if they are blank")
        
        # Compression moved under TOC checkbox in left column

        # Light separator above metadata section
        metadata_separator = tk.Frame(options_frame, height=1, bg="#D0D0D0")
        metadata_separator.pack(fill=tk.X, padx=8, pady=(6, 4))
        
        # Metadata section
        metadata_checkbox = tk.Checkbutton(
            options_frame,
            text="Add PDF metadata to combined file",
            variable=self.enable_metadata,
            command=self._toggle_metadata_fields,
            font=("Arial", 9)
        )
        metadata_checkbox.pack(anchor="w", padx=8, pady=(3, 1))
        
        # Title and Author on one line
        title_author_row = tk.Frame(options_frame)
        title_author_row.pack(fill=tk.X, pady=1, padx=8)
        tk.Label(title_author_row, text="Title:", font=("Arial", 9), width=10, anchor="e").pack(side=tk.LEFT)
        self.title_entry = tk.Entry(title_author_row, textvariable=self.pdf_title, font=("Arial", 9))
        self.title_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 10))
        self.title_entry.bind("<FocusOut>", lambda e: self._save_metadata_values())
        
        tk.Label(title_author_row, text="Author:", font=("Arial", 9), width=10, anchor="e").pack(side=tk.LEFT)
        self.author_entry = tk.Entry(title_author_row, textvariable=self.pdf_author, font=("Arial", 9))
        self.author_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        self.author_entry.bind("<FocusOut>", lambda e: self._save_metadata_values())
        
        # Subject and Keywords on one line
        subject_keywords_row = tk.Frame(options_frame)
        subject_keywords_row.pack(fill=tk.X, pady=1, padx=8)
        tk.Label(subject_keywords_row, text="Subject:", font=("Arial", 9), width=10, anchor="e").pack(side=tk.LEFT)
        self.subject_entry = tk.Entry(subject_keywords_row, textvariable=self.pdf_subject, font=("Arial", 9))
        self.subject_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 10))
        self.subject_entry.bind("<FocusOut>", lambda e: self._save_metadata_values())
        
        tk.Label(subject_keywords_row, text="Keywords:", font=("Arial", 9), width=10, anchor="e").pack(side=tk.LEFT)
        self.keywords_entry = tk.Entry(subject_keywords_row, textvariable=self.pdf_keywords, font=("Arial", 9))
        self.keywords_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        self.keywords_entry.bind("<FocusOut>", lambda e: self._save_metadata_values())
        
        # Initialize metadata field states
        self._toggle_metadata_fields()
        
        # Light separator above watermark section
        watermark_separator = tk.Frame(options_frame, height=1, bg="#D0D0D0")
        watermark_separator.pack(fill=tk.X, padx=8, pady=(6, 4))

        # Watermark section
        watermark_checkbox = tk.Checkbutton(
            options_frame,
            text="Add watermark to pages",
            variable=self.enable_watermark,
            command=self._toggle_watermark_fields,
            font=("Arial", 9)
        )
        watermark_checkbox.pack(anchor="w", padx=8, pady=(3, 1))
        
        # Watermark text and position on same row
        watermark_text_row = tk.Frame(options_frame)
        watermark_text_row.pack(fill=tk.X, pady=0, padx=8)
        tk.Label(watermark_text_row, text="Text:", font=("Arial", 9), width=10, anchor="e").pack(side=tk.LEFT)
        self.watermark_text_entry = tk.Entry(watermark_text_row, textvariable=self.watermark_text, font=("Arial", 9))
        self.watermark_text_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 10))
        self.watermark_text_entry.bind("<FocusOut>", lambda e: self._save_settings())
        
        tk.Label(watermark_text_row, text="Position:", font=("Arial", 9), width=10, anchor="e").pack(side=tk.LEFT)
        self.watermark_position_combo = ttk.Combobox(
            watermark_text_row,
            textvariable=self.watermark_position,
            values=["Top", "Center", "Bottom"],
            width=10,
            state="readonly",
            font=("Arial", 9)
        )
        self.watermark_position_combo.pack(side=tk.LEFT, padx=(5, 0))
        self.watermark_position_combo.bind("<<ComboboxSelected>>", lambda e: self._save_settings())
        
        # Opacity and Font Size on one line
        watermark_sliders_row = tk.Frame(options_frame)
        watermark_sliders_row.pack(fill=tk.X, pady=(1, 1), padx=8)
        
        tk.Label(watermark_sliders_row, text="Opacity:", font=("Arial", 9), width=10, anchor="e").pack(side=tk.LEFT)
        self.opacity_scale = tk.Scale(
            watermark_sliders_row,
            from_=0.1,
            to=0.9,
            resolution=0.1,
            orient=tk.HORIZONTAL,
            variable=self.watermark_opacity,
            showvalue=True,
            font=("Arial", 8),
            length=120,
            command=lambda e: self._save_settings()
        )
        self.opacity_scale.pack(side=tk.LEFT, padx=(5, 5))
        
        tk.Label(watermark_sliders_row, text="Font Size:", font=("Arial", 9), width=10, anchor="e").pack(side=tk.LEFT, padx=(10, 0))
        self.fontsize_scale = tk.Scale(
            watermark_sliders_row,
            from_=10,
            to=150,
            resolution=5,
            orient=tk.HORIZONTAL,
            variable=self.watermark_font_size,
            showvalue=True,
            font=("Arial", 8),
            length=120,
            command=lambda e: self._save_settings()
        )
        self.fontsize_scale.pack(side=tk.LEFT, padx=(5, 0))
        
        tk.Label(watermark_sliders_row, text="Rotation:", font=("Arial", 9), width=10, anchor="e").pack(side=tk.LEFT, padx=(10, 0))
        self.rotation_scale = tk.Scale(
            watermark_sliders_row,
            from_=0,
            to=360,
            resolution=5,
            orient=tk.HORIZONTAL,
            variable=self.watermark_rotation,
            showvalue=True,
            font=("Arial", 8),
            length=120,
            command=lambda e: self._save_settings()
        )
        self.rotation_scale.pack(side=tk.LEFT, padx=(5, 0))
        
        # Safe Mode checkbox
        watermark_safe_mode_row = tk.Frame(options_frame)
        watermark_safe_mode_row.pack(fill=tk.X, pady=(2, 2), padx=(28, 8))
        
        self.watermark_safe_mode_checkbox = tk.Checkbutton(
            watermark_safe_mode_row,
            text="Safe Mode: Auto-adjust watermark to attempt to prevent clipping",
            variable=self.watermark_safe_mode,
            command=self._save_settings,
            font=("Arial", 9)
        )
        self.watermark_safe_mode_checkbox.pack(side=tk.LEFT, anchor="w")
        ToolTip(self.watermark_safe_mode_checkbox, "When enabled, automatically reduces font size or adjusts\nposition to prevent watermark text from being clipped\nwhen combined with rotation and edge positioning.")
        
        # Initialize watermark field states
        self._toggle_watermark_fields()
        
        # ===== INSTRUCTIONS TAB =====
        instructions_frame = tk.Frame(self.notebook)
        self.notebook.add(instructions_frame, text="Instructions")
        
        # Load plain text instructions
        self._create_text_instructions_tab(instructions_frame)
        
        # ===== BOTTOM SECTION (Outside tabs) =====
        # Status bar frame
        status_frame = tk.Frame(root, bg="#E8E8E8", height=18)
        status_frame.pack(pady=0, padx=0, fill=tk.X, side=tk.BOTTOM)
        status_frame.pack_propagate(False)
        
        self.status_label = tk.Label(
            status_frame,
            text="",
            font=("Arial", 8),
            fg="#333333",
            bg="#E8E8E8",
            anchor="w",
            padx=5
        )
        self.status_label.pack(fill=tk.X, side=tk.LEFT)
        
        # Bottom button frame
        bottom_frame = tk.Frame(root, height=72)
        bottom_frame.pack(pady=5, padx=10, fill=tk.X)
        bottom_frame.pack_propagate(False)
        
        # Center frame for equal-width buttons
        center_frame = tk.Frame(bottom_frame)
        center_frame.place(relx=0.5, rely=0.5, anchor="center")
        
        # Combine button
        self.combine_button = tk.Button(
            center_frame,
            text="Combine PDFs",
            command=self.combine_pdfs,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 11),
            height=1,
            width=13,
            state=tk.DISABLED  # Start disabled until at least 2 files added
        )
        self.combine_button.pack(side=tk.LEFT, padx=5)
        
        # (Removed: open-button replaced by a post-success dialog)
        
        # Quit button
        quit_button = tk.Button(
            center_frame,
            text="Quit",
            command=self.root.quit,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 11),
            height=1,
            width=13
        )
        quit_button.pack(side=tk.LEFT, padx=5)

        # Left-side donate link
        left_info_frame = tk.Frame(bottom_frame, bg=bottom_frame.cget("bg"))
        left_info_frame.place(relx=0.0, rely=0.5, anchor="w", x=5)

        donate_label = tk.Label(
            left_info_frame,
            text="Like this? Donate!",
            font=("Arial", 10, "underline"),
            fg="#1A5FB4",
            bg=bottom_frame.cget("bg"),
            cursor="hand2"
        )
        donate_label.pack(anchor="w", pady=(0, 0))
        donate_label.bind(
            "<Button-1>",
            lambda e: webbrowser.open_new("https://www.paypal.com/paypalme/tgtechdevshop")
        )
        
        # Right-side stack for version and copyright link
        right_info_frame = tk.Frame(bottom_frame, bg=bottom_frame.cget("bg"))
        right_info_frame.place(relx=1.0, rely=0.5, anchor="e", x=-5)

        version_label = tk.Label(
            right_info_frame,
            text=f"v{__VERSION__}",
            font=("Arial", 8),
            fg="#606060",
            bg=bottom_frame.cget("bg")
        )
        version_label.pack(anchor="w", pady=(0, 0))

        copyright_label = tk.Label(
            right_info_frame,
            text="© 2026 tgtechy",
            font=("Arial", 8, "underline"),
            fg="#1A5FB4",
            bg=bottom_frame.cget("bg"),
            cursor="hand2"
        )
        copyright_label.pack(anchor="w", pady=(0, 0))
        copyright_label.bind(
            "<Button-1>",
            lambda e: webbrowser.open_new("https://github.com/tgtechy/CombinePDFs")
        )

        
        # Set up tab change handler to maintain focus on add button
        def on_tab_changed(event):
            if self.notebook.index(self.notebook.select()) == 0:  # Input tab
                self.add_button.focus_set()
        
        self.notebook.bind("<<NotebookTabChanged>>", on_tab_changed)
        
        # Set initial focus to add button
        self.add_button.focus_set()
        
        # Initialize the file list display (shows placeholder if empty)
        self.refresh_listbox()
    
    # Helper methods for file dict access
    def get_file_path(self, file_entry: dict) -> str:
        """Extract file path from file entry dict"""
        return file_entry['path']
    
    def get_rotation(self, file_entry: dict) -> int:
        """Extract rotation value from file entry dict"""
        return file_entry.get('rotation', 0)
    
    def set_rotation(self, index: int, degrees: int):
        """Update rotation for a file at given index"""
        if 0 <= index < len(self.pdf_files):
            self.pdf_files[index]['rotation'] = degrees
            self.refresh_listbox()
    
    def get_reverse(self, file_entry: dict) -> bool:
        """Extract reverse value from file entry dict"""
        return file_entry.get('reverse', False)
    
    def set_reverse(self, index: int, reverse: bool):
        """Update reverse setting for a file at given index"""
        if 0 <= index < len(self.pdf_files):
            self.pdf_files[index]['reverse'] = reverse
    
    def get_page_range(self, file_entry: dict) -> str:
        """Extract page range from file entry dict"""
        return file_entry.get('page_range', 'All')
    
    def set_page_range(self, index: int, page_range: str):
        """Update page range for a file at given index"""
        if 0 <= index < len(self.pdf_files):
            cleaned = page_range.strip()
            self.pdf_files[index]['page_range'] = cleaned if cleaned else 'All'
    
    def _load_settings(self):
        """Load saved settings from config file"""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r') as f:
                    settings = json.load(f)
                    
                # Load output directory if it exists
                if 'output_directory' in settings:
                    saved_dir = settings['output_directory']
                    if os.path.exists(saved_dir):
                        self.output_directory = saved_dir
                
                # Load add files directory if it exists
                if 'add_files_directory' in settings:
                    saved_dir = settings['add_files_directory']
                    if os.path.exists(saved_dir):
                        self.add_files_directory = saved_dir
                
                # Load list files directory if it exists
                if 'list_files_directory' in settings:
                    saved_dir = settings['list_files_directory']
                    if os.path.exists(saved_dir):
                        self.list_files_directory = saved_dir
                
                # Load preview enabled state
                if 'preview_enabled' in settings:
                    self.preview_enabled.set(settings['preview_enabled'])
                
                # Load add filename bookmarks state
                if 'add_filename_bookmarks' in settings:
                    self.add_filename_bookmarks.set(settings['add_filename_bookmarks'])
                
                # Load insert blank pages state
                if 'insert_blank_pages' in settings:
                    self.insert_blank_pages.set(settings['insert_blank_pages'])
                
                # Load breaker pages uniform size state
                if 'breaker_pages_uniform_size' in settings:
                    self.breaker_pages_uniform_size.set(settings['breaker_pages_uniform_size'])
                
                # Load insert TOC state
                if 'insert_toc' in settings:
                    self.insert_toc.set(settings['insert_toc'])
                
                # Load advanced settings
                if 'compression_quality' in settings:
                    self.compression_quality.set(settings['compression_quality'])
                
                # Load last used metadata values
                if 'last_metadata' in settings:
                    self.last_metadata = settings['last_metadata']
                else:
                    # Initialize with saved values if they exist
                    self.last_metadata = {
                        'title': settings.get('pdf_title', ''),
                        'author': settings.get('pdf_author', ''),
                        'subject': settings.get('pdf_subject', ''),
                        'keywords': settings.get('pdf_keywords', '')
                    }
                
                # Load metadata fields (will be cleared if metadata not enabled)
                if 'pdf_title' in settings:
                    self.pdf_title.set(settings['pdf_title'])
                if 'pdf_author' in settings:
                    self.pdf_author.set(settings['pdf_author'])
                if 'pdf_subject' in settings:
                    self.pdf_subject.set(settings['pdf_subject'])
                if 'pdf_keywords' in settings:
                    self.pdf_keywords.set(settings['pdf_keywords'])
                if 'enable_metadata' in settings:
                    self.enable_metadata.set(settings['enable_metadata'])
                if 'enable_page_scaling' in settings:
                    self.enable_page_scaling.set(settings['enable_page_scaling'])
                if 'enable_watermark' in settings:
                    self.enable_watermark.set(settings['enable_watermark'])
                if 'watermark_text' in settings:
                    self.watermark_text.set(settings['watermark_text'])
                if 'watermark_opacity' in settings:
                    self.watermark_opacity.set(settings['watermark_opacity'])
                if 'watermark_font_size' in settings:
                    self.watermark_font_size.set(settings['watermark_font_size'])
                if 'watermark_rotation' in settings:
                    self.watermark_rotation.set(settings['watermark_rotation'])
                if 'watermark_position' in settings:
                    self.watermark_position.set(settings['watermark_position'])
                if 'watermark_safe_mode' in settings:
                    self.watermark_safe_mode.set(settings['watermark_safe_mode'])
                if 'delete_blank_pages' in settings:
                    self.delete_blank_pages.set(settings['delete_blank_pages'])
        except Exception:
            # If loading fails, just use defaults
            pass
    
    def _save_settings(self):
        """Save current settings to config file"""
        try:
            # Ensure config directory exists
            self.config_file.parent.mkdir(parents=True, exist_ok=True)
            
            settings = {
                'output_directory': self.output_directory,
                'add_files_directory': self.add_files_directory,
                'list_files_directory': self.list_files_directory,
                'preview_enabled': self.preview_enabled.get(),
                'add_filename_bookmarks': self.add_filename_bookmarks.get(),
                'insert_blank_pages': self.insert_blank_pages.get(),
                'breaker_pages_uniform_size': self.breaker_pages_uniform_size.get(),
                'insert_toc': self.insert_toc.get(),
                'compression_quality': self.compression_quality.get(),
                'last_metadata': self.last_metadata,
                'pdf_title': self.pdf_title.get(),
                'pdf_author': self.pdf_author.get(),
                'pdf_subject': self.pdf_subject.get(),
                'pdf_keywords': self.pdf_keywords.get(),
                'enable_metadata': self.enable_metadata.get(),
                'enable_page_scaling': self.enable_page_scaling.get(),
                'enable_watermark': self.enable_watermark.get(),
                'watermark_text': self.watermark_text.get(),
                'watermark_opacity': self.watermark_opacity.get(),
                'watermark_font_size': self.watermark_font_size.get(),
                'watermark_rotation': self.watermark_rotation.get(),
                'watermark_position': self.watermark_position.get(),
                'watermark_safe_mode': self.watermark_safe_mode.get(),
                'delete_blank_pages': self.delete_blank_pages.get()
            }
            with open(self.config_file, 'w') as f:
                json.dump(settings, f, indent=2)
        except Exception:
            # Silently fail if we can't save settings
            pass
    
    def _toggle_breaker_page_options(self):
        """Enable or disable breaker page suboptions based on checkbox state"""
        state = tk.NORMAL if self.insert_blank_pages.get() else tk.DISABLED
        self.breaker_uniform_checkbox.config(state=state)
        self._save_settings()
    
    def _toggle_metadata_fields(self):
        """Enable or disable metadata entry fields based on checkbox state"""
        state = tk.NORMAL if self.enable_metadata.get() else tk.DISABLED
        self.title_entry.config(state=state)
        self.author_entry.config(state=state)
        self.subject_entry.config(state=state)
        self.keywords_entry.config(state=state)
        
        if self.enable_metadata.get():
            # Restore last used metadata values
            self.pdf_title.set(self.last_metadata.get('title', ''))
            self.pdf_author.set(self.last_metadata.get('author', ''))
            self.pdf_subject.set(self.last_metadata.get('subject', ''))
            self.pdf_keywords.set(self.last_metadata.get('keywords', ''))
            
            # If author is empty, populate with current username
            if not self.pdf_author.get():
                import getpass
                try:
                    username = getpass.getuser()
                    self.pdf_author.set(username)
                except Exception:
                    pass
        else:
            # Save current values to last_metadata (keep displaying but grayed out)
            self.last_metadata = {
                'title': self.pdf_title.get(),
                'author': self.pdf_author.get(),
                'subject': self.pdf_subject.get(),
                'keywords': self.pdf_keywords.get()
            }
        
        self._save_settings()
    
    def _validate_compression_quality(self):
        """Ensure compression quality always has a valid value"""
        valid_values = ["None", "Low", "Medium", "High", "Maximum"]
        current = self.compression_quality.get()
        if current not in valid_values:
            # Reset to default if invalid or empty
            self.compression_quality.set("Medium")
            self._save_settings()
    
    def _save_metadata_values(self):
        """Update last_metadata with current field values and save settings"""
        self.last_metadata = {
            'title': self.pdf_title.get(),
            'author': self.pdf_author.get(),
            'subject': self.pdf_subject.get(),
            'keywords': self.pdf_keywords.get()
        }
        self._save_settings()
    
    def _validate_rotation(self, index: int, var: tk.StringVar):
        """Ensure rotation dropdown always has a valid value"""
        valid_values = ["0", "90", "180", "270"]
        current = var.get()
        if current not in valid_values:
            # Reset to default if invalid or empty
            var.set("0")
            self.set_rotation(index, 0)
    

    def _validate_filename_on_focus_out(self):
        """Validate filename when the input box loses focus"""
        filename = self.output_filename.get().strip()
        if not filename:
            # Empty is okay, don't show error on blur
            return
        
        is_valid, error_message, corrected_filename = self._validate_output_filename(filename)
        if not is_valid:
            # For invalid characters, auto-correct and notify
            if "invalid characters" in error_message.lower():
                messagebox.showwarning("Filename Correction", error_message)
                self.output_filename.set(corrected_filename)
            else:
                # For other issues, just show the error
                self.show_error_dialog("Invalid Filename", error_message)
    
    def _toggle_watermark_fields(self):
        """Enable or disable watermark entry fields and sliders based on checkbox state"""
        state = tk.NORMAL if self.enable_watermark.get() else tk.DISABLED
        self.watermark_text_entry.config(state=state)
        self.opacity_scale.config(state=state)
        self.fontsize_scale.config(state=state)
        self.rotation_scale.config(state=state)
        self.watermark_safe_mode_checkbox.config(state=state)
        
        self._save_settings()
    
    def add_files(self):
        """Open file dialog to select PDF and image files"""
        files = filedialog.askopenfilenames(
            title="Select PDF and image files to combine",
            filetypes=[
                ("PDF and Image files", "*.pdf *.jpg *.jpeg *.png *.bmp *.gif *.tiff *.tif"),
                ("PDF files", "*.pdf"),
                ("Image files", "*.jpg *.jpeg *.png *.bmp *.gif *.tiff *.tif"),
                ("All files", "*.*")
            ],
            initialdir=self.add_files_directory
        )
        
        # Update the add files directory for next time if files were selected
        if files:
            self.add_files_directory = str(Path(files[0]).parent)
            self._save_settings()
        
        added_count = 0
        duplicate_count = 0
        duplicates = []
        unsupported_count = 0
        unsupported_files = []
        
        # Get existing paths for duplicate checking
        existing_paths = {entry['path'] for entry in self.pdf_files}
        
        # Supported image and PDF extensions
        supported_exts = ('.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.tif')
        
        for file in files:
            # Check if file has a supported extension
            if not file.lower().endswith(supported_exts):
                unsupported_count += 1
                unsupported_files.append(Path(file).name)
                continue
            
            if file not in existing_paths:
                # Create dict entry with path and default rotation/page range/reverse
                # Note: rotation applies to PDFs; images will be rotated after conversion
                self.pdf_files.append({'path': file, 'rotation': 0, 'page_range': 'All', 'reverse': False})
                added_count += 1
            else:
                duplicate_count += 1
                duplicates.append(Path(file).name)

        # Refresh list display with updated numbering
        try:
            # Clear any active sort when files are added
            self.sort_key = None
            self.sort_reverse = False
            self.refresh_listbox()
            self.update_header_labels()
        except Exception:
            self.refresh_listbox()

        # Clear status bar
        self.status_label.config(text="")

        self.update_count()
        
        # Show warning if unsupported files were attempted
        if unsupported_count > 0:
            unsupported_text = "\n".join(f"  • {file}" for file in unsupported_files)
            messagebox.showwarning(
                "Unsupported Files",
                f"The following file(s) are not PDF or image files and were not added:\n\n{unsupported_text}"
            )
        
        # Show warning if duplicates were attempted
        if duplicate_count > 0:
            duplicates_text = "\n".join(f"  • {dup}" for dup in duplicates)
            messagebox.showwarning(
                "Duplicate Files",
                f"The following file(s) are already in the list and were not added:\n\n{duplicates_text}"
            )
    
    def get_file_path(self, file_entry: Dict[str, any]) -> str:
        """Extract file path from entry dict"""
        return file_entry['path']

    def get_rotation(self, file_entry: Dict[str, any]) -> int:
        """Extract rotation value from entry dict"""
        return file_entry.get('rotation', 0)

    def get_reverse(self, file_entry: Dict[str, any]) -> bool:
        """Extract reverse value from entry dict"""
        return file_entry.get('reverse', False)

    def get_page_range(self, file_entry: Dict[str, any]) -> str:
        """Extract page range from entry dict"""
        return file_entry.get('page_range', 'All')
    
    def remove_file(self):
        """Remove selected file(s) from list"""
        try:
            # Find selected rows by checking which ones have selection highlighting
            selected_indices = []
            rows = self.file_list_frame.winfo_children()
            for i, row in enumerate(rows):
                if hasattr(row, '_is_selected') and row._is_selected:
                    selected_indices.append(i)
            
            if not selected_indices:
                messagebox.showwarning("Warning", "Please select a file to remove from the list.")
                return

            # Confirm removal
            count = len(selected_indices)
            file_word = "file" if count == 1 else "files"
            if not messagebox.askyesno("Confirm Removal", f"Remove {count} selected {file_word} from the list?"):
                return

            # Delete in reverse order to avoid index shifting
            for index in reversed(selected_indices):
                del self.pdf_files[index]

            # Clear status bar
            self.status_label.config(text="")
            
            # Refresh display and count; clear sort state so arrows disappear
            try:
                self.sort_key = None
                self.sort_reverse = False
                self.refresh_listbox()
                self.update_header_labels()
            except Exception:
                self.refresh_listbox()

            self.update_count()
        except Exception:
            messagebox.showwarning("Warning", "Please select a file to remove from the list.")
    
    def clear_files(self):
        """Clear all files from list"""
        if not self.pdf_files:
            return
        
        count = len(self.pdf_files)
        file_word = "file" if count == 1 else "files"
        if not messagebox.askyesno("Confirm Clear All", f"Remove all {count} {file_word} from the list?"):
            return
        
        self.pdf_files.clear()
        self.rotation_vars.clear()
        self.page_range_vars.clear()
        self.page_range_last_valid.clear()
        
        # Clear status bar
        self.status_label.config(text="")
        
        try:
            # Clear any active sort when list is cleared
            self.sort_key = None
            self.sort_reverse = False
            self.refresh_listbox()
            self.update_header_labels()
        except Exception:
            self.refresh_listbox()

        self.update_count()
    
    def update_count(self):
        """Update the file count label"""
        count = len(self.pdf_files)
        self.count_label.config(text=f"Files to combine: {count}")
        # Also update button states when file count changes
        self._update_button_states()
    
    def show_load_save_dialog(self):
        """Show a dialog window for loading or saving PDF lists"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Load/Save PDF List")
        dialog_width, dialog_height = self._scale_geometry(400, 300)
        dialog.geometry(f"{dialog_width}x{dialog_height}")
        dialog.resizable(False, False)
        dialog.withdraw()  # Hide initially to position before showing
        
        # Make dialog modal
        dialog.transient(self.root)
        
        # Title
        title_label = tk.Label(dialog, text="PDF List Management", font=("Arial", 12, "bold"))
        title_label.pack(pady=(20, 10))
        
        # Description text
        description_label = tk.Label(
            dialog,
            text="Save your current PDF list to reuse later,\nor load a previously saved list.",
            font=("Arial", 9),
            fg="#666666"
        )
        description_label.pack(pady=(0, 20))
        
        # Button frame
        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=20, padx=20, fill=tk.BOTH, expand=True)
        
        # Save list button
        save_button = tk.Button(
            button_frame,
            text="Save Current List",
            command=lambda: self.save_pdf_list(dialog),
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10),
            height=4
        )
        save_button.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        # Load list button
        load_button = tk.Button(
            button_frame,
            text="Load Previously Saved List",
            command=lambda: self.load_pdf_list(dialog),
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10),
            height=4
        )
        load_button.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        
        # Close button
        close_button = tk.Button(
            button_frame,
            text="Close",
            command=dialog.destroy,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10),
            height=4
        )
        close_button.grid(row=2, column=0, sticky="nsew", padx=5, pady=5)
        
        # Configure grid to make buttons equal width and height
        button_frame.columnconfigure(0, weight=1)
        button_frame.rowconfigure(0, weight=1)
        button_frame.rowconfigure(1, weight=1)
        button_frame.rowconfigure(2, weight=1)
        
        # Update to get accurate dimensions
        dialog.update_idletasks()
        
        # Get parent window position and size
        parent_x = self.root.winfo_x()
        parent_y = self.root.winfo_y()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()
        
        # Get dialog size
        dialog_width = dialog.winfo_width()
        dialog_height = dialog.winfo_height()
        
        # Calculate center position relative to parent window
        center_x = parent_x + (parent_width - dialog_width) // 2
        center_y = parent_y + (parent_height - dialog_height) // 2
        
        dialog.geometry(f"+{center_x}+{center_y}")
        
        # Show dialog
        dialog.deiconify()
        dialog.grab_set()
    
    def save_pdf_list(self, dialog):
        """Save the current list of PDFs to a JSON file"""
        if not self.pdf_files:
            messagebox.showwarning("Empty List", "There are no PDFs to save. Add some files first.")
            return
        
        # Ask user for file location
        file_path = filedialog.asksaveasfilename(
            initialdir=self.list_files_directory,
            filetypes=[("PDF List Files", "*.pdflist"), ("JSON Files", "*.json"), ("All Files", "*.*")],
            defaultextension=".pdflist"
        )
        
        if not file_path:
            return
        
        # Update the list files directory for next time
        self.list_files_directory = str(Path(file_path).parent)
        self._save_settings()
        
        try:
            # Convert pdf_files to a serializable format
            list_data = []
            for pdf_entry in self.pdf_files:
                list_data.append({
                    'path': self.get_file_path(pdf_entry),
                    'rotation': self.get_rotation(pdf_entry),
                    'page_range': self.get_page_range(pdf_entry),
                    'reverse': self.get_reverse(pdf_entry)
                })
            
            # Write to JSON file
            with open(file_path, 'w') as f:
                json.dump(list_data, f, indent=2)
            
            self.show_info_dialog("Success", f"PDF list saved to:\n{file_path}")
            dialog.destroy()
        except Exception as e:
            self.show_error_dialog("Error", f"Failed to save list:\n{str(e)}")
    
    def load_pdf_list(self, dialog):
        """Load a previously saved PDF list from a JSON file"""
        # Ask user for file location
        file_path = filedialog.askopenfilename(
            initialdir=self.list_files_directory,
            filetypes=[("PDF List Files", "*.pdflist"), ("JSON Files", "*.json"), ("All Files", "*.*")]
        )
        
        if not file_path:
            return
        
        # Update the list files directory for next time
        self.list_files_directory = str(Path(file_path).parent)
        self._save_settings()
        
        try:
            # Read JSON file
            with open(file_path, 'r') as f:
                list_data = json.load(f)
            
            # Verify and create file entries
            valid_entries = []
            missing_files = []
            
            for entry in list_data:
                file_path_to_check = entry.get('path')
                if os.path.exists(file_path_to_check):
                    valid_entries.append({
                        'path': file_path_to_check,
                        'rotation': entry.get('rotation', 0),
                        'page_range': entry.get('page_range', 'All'),
                        'reverse': entry.get('reverse', False)
                    })
                else:
                    missing_files.append(os.path.basename(file_path_to_check))
            
            if not valid_entries:
                self.show_error_dialog("Error", "No valid PDF files found in the saved list.")
                return
            
            # Ask if user wants to replace or append
            if self.pdf_files:
                result = self.show_merge_lists_dialog(len(valid_entries))
                if result is None:
                    return
                elif not result:  # Replace
                    self.pdf_files.clear()
            
            # Add the loaded entries, filtering out duplicates if appending
            existing_paths = {entry['path'] for entry in self.pdf_files}
            new_entries = [entry for entry in valid_entries if entry['path'] not in existing_paths]
            duplicate_count = len(valid_entries) - len(new_entries)
            
            self.pdf_files.extend(new_entries)
            self.refresh_listbox()
            self.update_count()
            
            # Show message about missing files if any
            if missing_files:
                missing_text = "\n".join(missing_files)
                messagebox.showwarning(
                    "Missing Files",
                    f"The following files were not found and were skipped:\n\n{missing_text}"
                )
            else:
                success_msg = f"Loaded {len(new_entries)} PDF files from the saved list."
                if duplicate_count > 0:
                    dup_word = "file" if duplicate_count == 1 else "files"
                    success_msg += f"\n({duplicate_count} duplicate {dup_word} skipped)"
                self.show_info_dialog("Success", success_msg)
            
            dialog.destroy()
        except json.JSONDecodeError:
            self.show_error_dialog("Error", "The file is not a valid PDF list file.")
        except Exception as e:
            self.show_error_dialog("Error", f"Failed to load list:\n{str(e)}")
    
    def get_file_info(self, file_path: str) -> tuple:
        """Get formatted file info. Returns tuple of (filename, filesize_str, date_str)"""
        try:
            file_stat = os.stat(file_path)
            size_bytes = file_stat.st_size
            
            # Format file size
            if size_bytes < 1024:
                size_str = f"{size_bytes} B"
            elif size_bytes < 1024 * 1024:
                size_str = f"{size_bytes / 1024:.1f} KB"
            else:
                size_str = f"{size_bytes / (1024 * 1024):.1f} MB"
            
            # Format modification date
            mod_time = datetime.fromtimestamp(file_stat.st_mtime)
            date_str = mod_time.strftime("%m/%d/%Y")
            
            filename = Path(file_path).name
            # Truncate long filenames so columns remain aligned
            max_filename_len = 60
            if len(filename) > max_filename_len:
                filename = filename[: max_filename_len - 3] + "..."

            return (filename, size_str, date_str)
        except Exception:
            return (Path(file_path).name, "N/A", "N/A")

    def format_list_item(self, index: int, file_entry: Dict[str, any]) -> str:
        """Return formatted string for display. No longer used with custom frame, but kept for reference."""
        file_path = self.get_file_path(file_entry)
        filename, size_str, date_str = self.get_file_info(file_path)
        rotation = self.get_rotation(file_entry)
        return f"{index+1:>3}. {filename:<55} {size_str:>12}  {date_str}  {rotation}°"

    def refresh_listbox(self):
        """Rebuild the custom list frame from `self.pdf_files` with rotation controls."""
        # Clear existing rows
        for widget in self.file_list_frame.winfo_children():
            widget.destroy()
        
        self.rotation_vars.clear()
        self.page_range_vars.clear()
        self.reverse_vars.clear()
        self.row_visual_state.clear()  # Clear cached visual state since rows are rebuilt
        
        # Update button states after clearing
        self.root.after_idle(self._update_button_states)
        
        # Show placeholder if list is empty
        if len(self.pdf_files) == 0:
            # Get canvas height for placeholder centering
            canvas_height = self.file_list_canvas.winfo_height()
            if canvas_height <= 1:
                canvas_height = 270
            
            # Set frame height to canvas height so placeholder can center vertically
            self.file_list_frame.configure(height=canvas_height)
            
            placeholder_frame = tk.Frame(self.file_list_frame, bg="white")
            placeholder_frame.pack(fill=tk.BOTH, expand=True)
            
            placeholder_label = tk.Label(
                placeholder_frame,
                text='Click the "Add PDFs/Images" button below to get started\nSupported formats are PDF, JPG, PNG, BMP, GIF, TIFF\n\nClick the tabs at the top to switch between\ninput file selection and output (combining) settings',
                font=("Arial", 11, "bold"),
                fg="red",
                bg="white",
                justify=tk.CENTER,
                anchor="center"
            )
            placeholder_label.pack(fill=tk.BOTH, expand=True)
            
            # Manually update canvas scrollregion after adding placeholder
            self.root.after_idle(self.canvas_configure)
            return
        
        # Reset frame height to natural size when files are present
        self.file_list_frame.configure(height=1)
        
        for i, pdf_entry in enumerate(self.pdf_files):
            file_path = self.get_file_path(pdf_entry)
            rotation = self.get_rotation(pdf_entry)
            is_image = file_path.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.tif'))
            row_fg = "#0052CC" if is_image else "black"
            
            # Create row frame - disable focus to prevent focus-change flicker
            row_frame = tk.Frame(self.file_list_frame, bg="white", takefocus=0, height=20)
            row_frame.pack(fill=tk.X, padx=0, pady=0, anchor='nw')
            row_frame.pack_propagate(False)  # Maintain fixed height
            row_frame._index = i
            row_frame._is_selected = False
            
            # Register drag and drop events on row
            row_frame.bind("<Button-1>", lambda e, idx=i: self.on_row_click(e, idx))
            row_frame.bind("<B1-Motion>", lambda e, idx=i: self.on_row_drag(e, idx))
            row_frame.bind("<ButtonRelease-1>", lambda e, idx=i: self.on_row_release(e, idx))
            row_frame.bind("<Motion>", lambda e, idx=i: self.on_row_hover(e, idx))
            row_frame.bind("<Leave>", self.on_row_leave)
            row_frame.bind("<Double-Button-1>", lambda e, idx=i: self.on_row_double_click(e, idx))
            
            # Get file info
            filename, size_str, date_str = self.get_file_info(file_path)
            
            # Number label
            num_label = tk.Label(row_frame, text=f"{i+1}", font=("Consolas", 9), bg="white", fg=row_fg, width=4, anchor='e')
            num_label.pack(side=tk.LEFT, padx=(0, 2), pady=0, ipady=0, anchor='nw')
            num_label.bind("<Button-1>", lambda e, idx=i: self.on_row_click(e, idx))
            num_label.bind("<B1-Motion>", lambda e, idx=i: self.on_row_drag(e, idx))
            num_label.bind("<ButtonRelease-1>", lambda e, idx=i: self.on_row_release(e, idx))
            num_label.bind("<Motion>", lambda e, idx=i: self.on_row_hover(e, idx))
            num_label.bind("<Leave>", self.on_row_leave)
            num_label.bind("<Double-Button-1>", lambda e, idx=i: self.on_row_double_click(e, idx))
            
            # Filename label
            filename_label = tk.Label(row_frame, text=filename, font=("Consolas", 9), bg="white", fg=row_fg, width=62, anchor='w', justify=tk.LEFT)
            filename_label.pack(side=tk.LEFT, pady=0, ipady=0, anchor='nw')
            filename_label.bind("<Button-1>", lambda e, idx=i: self.on_row_click(e, idx))
            filename_label.bind("<B1-Motion>", lambda e, idx=i: self.on_row_drag(e, idx))
            filename_label.bind("<ButtonRelease-1>", lambda e, idx=i: self.on_row_release(e, idx))
            filename_label.bind("<Motion>", lambda e, idx=i: self.on_row_hover(e, idx))
            filename_label.bind("<Leave>", self.on_row_leave)
            filename_label.bind("<Double-Button-1>", lambda e, idx=i: self.on_row_double_click(e, idx))
            
            # File size label
            size_label = tk.Label(row_frame, text=size_str, font=("Consolas", 9), bg="white", fg=row_fg, width=10, anchor='w')
            size_label.pack(side=tk.LEFT, pady=0, ipady=0, anchor='nw')
            size_label.bind("<Button-1>", lambda e, idx=i: self.on_row_click(e, idx))
            size_label.bind("<B1-Motion>", lambda e, idx=i: self.on_row_drag(e, idx))
            size_label.bind("<ButtonRelease-1>", lambda e, idx=i: self.on_row_release(e, idx))
            size_label.bind("<Motion>", lambda e, idx=i: self.on_row_hover(e, idx))
            size_label.bind("<Leave>", self.on_row_leave)
            size_label.bind("<Double-Button-1>", lambda e, idx=i: self.on_row_double_click(e, idx))
            
            # Date label
            date_label = tk.Label(row_frame, text=date_str, font=("Consolas", 9), bg="white", fg=row_fg, width=11, anchor='w')
            date_label.pack(side=tk.LEFT, pady=0, ipady=0, anchor='nw')
            date_label.bind("<Button-1>", lambda e, idx=i: self.on_row_click(e, idx))
            date_label.bind("<B1-Motion>", lambda e, idx=i: self.on_row_drag(e, idx))
            date_label.bind("<ButtonRelease-1>", lambda e, idx=i: self.on_row_release(e, idx))
            date_label.bind("<Motion>", lambda e, idx=i: self.on_row_hover(e, idx))
            date_label.bind("<Leave>", self.on_row_leave)
            date_label.bind("<Double-Button-1>", lambda e, idx=i: self.on_row_double_click(e, idx))
            
            # Page range entry
            page_range = self.get_page_range(pdf_entry)
            page_range_var = tk.StringVar(value=page_range)
            self.page_range_vars[i] = page_range_var
            self.page_range_last_valid[i] = page_range
            
            page_entry = tk.Entry(
                row_frame,
                textvariable=page_range_var,
                width=7,
                font=("Consolas", 9),
                state="disabled" if is_image else "normal",
                disabledforeground="#999999" if is_image else "black",
                disabledbackground="#F5F5F5" if is_image else "white"
            )
            page_entry.pack(side=tk.LEFT, padx=(4, 0), pady=0, ipady=0, anchor='nw')
            
            def on_page_range_change(var, idx=i):
                self.set_page_range(idx, var.get())
            
            page_range_var.trace("w", lambda *args, var=page_range_var, idx=i: on_page_range_change(var, idx))
            page_entry.bind("<FocusOut>", lambda e, idx=i, var=page_range_var, ent=page_entry: self._validate_page_range(idx, var, ent))
            page_entry.bind("<Return>", lambda e, idx=i, var=page_range_var, ent=page_entry: self._validate_page_range(idx, var, ent))
            
            page_entry.bind("<Button-1>", lambda e, idx=i: self.on_row_click(e, idx))
            page_entry.bind("<Motion>", lambda e, idx=i: self.on_row_hover(e, idx))
            page_entry.bind("<Leave>", self.on_row_leave)
            
            # Rotation dropdown
            rotation_var = tk.StringVar(value=f"{rotation}°")
            self.rotation_vars[i] = rotation_var
            
            rotation_dropdown = ttk.Combobox(
                row_frame,
                textvariable=rotation_var,
                values=["0°", "90°", "180°", "270°"],
                width=4,
                state="readonly",
                font=("Consolas", 9)
            )
            rotation_dropdown.pack(side=tk.LEFT, padx=2, pady=0, ipady=0, anchor='nw')
            
            # Bind rotation change
            def on_rotation_change(var, idx=i):
                try:
                    value = var.get()
                    if value.endswith("°"):
                        value = value[:-1]
                    degrees = int(value)
                    self.set_rotation(idx, degrees)
                except ValueError:
                    pass
            
            rotation_var.trace("w", lambda *args, var=rotation_var, idx=i: on_rotation_change(var, idx))
            
            # Bind FocusOut to validate rotation value
            rotation_dropdown.bind("<FocusOut>", lambda e, idx=i, var=rotation_var: self._validate_rotation(idx, var))
            
            # Bind events to dropdown too for consistency
            rotation_dropdown.bind("<Button-1>", lambda e, idx=i: self.on_row_click(e, idx))
            rotation_dropdown.bind("<Motion>", lambda e, idx=i: self.on_row_hover(e, idx))
            rotation_dropdown.bind("<Leave>", self.on_row_leave)
            
            # Reverse pages checkbox
            reverse = self.get_reverse(pdf_entry)
            reverse_var = tk.BooleanVar(value=reverse)
            self.reverse_vars[i] = reverse_var
            
            reverse_checkbox = tk.Checkbutton(
                row_frame,
                variable=reverse_var,
                command=lambda idx=i, var=reverse_var: self.set_reverse(idx, var.get()),
                bg="white",
                takefocus=0,
                state="disabled" if is_image else "normal",
                disabledforeground="#999999" if is_image else "black"
            )
            reverse_checkbox.pack(side=tk.LEFT, padx=2, pady=0, anchor='nw')
            
            # Bind events to checkbox for consistency
            reverse_checkbox.bind("<Button-1>", lambda e, idx=i: self.on_row_click(e, idx), add="+")
            reverse_checkbox.bind("<Motion>", lambda e, idx=i: self.on_row_hover(e, idx))
            reverse_checkbox.bind("<Leave>", self.on_row_leave)
        
        # Manually update canvas scrollregion after rebuilding list
        self.root.after_idle(self.canvas_configure)
    
    def on_row_click(self, event, index: int):
        """Handle mouse down event for drag and drop"""
        # Find the row frame
        rows = self.file_list_frame.winfo_children()
        if index < len(rows):
            row_frame = rows[index]
            
            # Toggle selection
            if event.state & 0x0004:  # Ctrl key
                row_frame._is_selected = not row_frame._is_selected
            else:
                # Clear all other selections
                for row in rows:
                    row._is_selected = False
                row_frame._is_selected = True
            
            # Update status bar with selected file path
            if 0 <= index < len(self.pdf_files):
                file_entry = self.pdf_files[index]
                file_path = self.get_file_path(file_entry)
                self.status_label.config(text=file_path)
            
            # Update visuals immediately
            self._update_row_visuals()
            self.drag_start_index = index
            self.drag_start_y = event.y_root
            self.is_dragging = False
    
    def on_row_drag(self, event, index: int):
        """Handle mouse drag event"""
        if self.drag_start_index is None:
            return
        
        # Only start dragging if mouse has moved more than 5 pixels
        if not self.is_dragging:
            if abs(event.y_root - self.drag_start_y) < 5:
                return  # Not enough movement to constitute a drag
            self.is_dragging = True
        
        # Auto-scroll when dragging near edges of canvas
        self._auto_scroll_during_drag(event)
        
        # Get current position
        current_y = event.y_root
        rows = self.file_list_frame.winfo_children()
        
        # Find which row we're over by converting to coordinates relative to file_list_frame
        drag_y = self.file_list_frame.winfo_pointery() - self.file_list_frame.winfo_rooty()
        
        # Find target index
        current_index = None
        for i, row in enumerate(rows):
            row_y = row.winfo_y()
            row_height = row.winfo_height()
            if drag_y >= row_y - row_height // 2 and drag_y < row_y + row_height // 2:
                current_index = i
                break
        
        if current_index is not None and current_index != self.drag_start_index and 0 <= current_index < len(self.pdf_files):
            # Reorder the backing list
            dragged_entry = self.pdf_files.pop(self.drag_start_index)
            self.pdf_files.insert(current_index, dragged_entry)
            self.drag_start_index = current_index

            # Refresh and restore selection
            self.refresh_listbox()
            rows = self.file_list_frame.winfo_children()
            if current_index < len(rows):
                rows[current_index]._is_selected = True
            self._update_row_visuals()

            # Clear sort indicators
            self.sort_key = None
            self.sort_reverse = False
            self.update_header_labels()
    
    def on_row_release(self, event, index: int):
        """Handle mouse up event"""
        self.is_dragging = False
        self.drag_start_index = None
        self.drag_start_y = None
        # Cancel any pending auto-scroll immediately and thoroughly
        if self.auto_scroll_id:
            try:
                self.root.after_cancel(self.auto_scroll_id)
            except Exception:
                pass
            self.auto_scroll_id = None
    
    def _auto_scroll_during_drag(self, event):
        """Auto-scroll the canvas when dragging near top or bottom edges"""
        # Exit immediately if dragging stopped
        if not self.is_dragging:
            self.auto_scroll_id = None
            return
        
        try:
            # Get mouse position relative to canvas
            canvas_y = event.y_root - self.file_list_canvas.winfo_rooty()
            canvas_height = self.file_list_canvas.winfo_height()
        except Exception:
            # Event might be invalid, stop scrolling
            self.auto_scroll_id = None
            return
        
        scroll_zone = 30  # Pixels from edge to trigger scrolling
        scroll_speed = 1  # Lines to scroll per update
        
        # Check if near top or bottom
        if canvas_y < scroll_zone and self.is_dragging:
            # Near top - scroll up (only if not already at top)
            first, last = self.file_list_canvas.yview()
            if first <= 0:
                if self.auto_scroll_id:
                    try:
                        self.root.after_cancel(self.auto_scroll_id)
                    except Exception:
                        pass
                    self.auto_scroll_id = None
                return
            self.file_list_canvas.yview_scroll(-1, "units")
            # Schedule next scroll only if still dragging
            if self.auto_scroll_id and self.is_dragging:
                try:
                    self.root.after_cancel(self.auto_scroll_id)
                except Exception:
                    pass
            self.auto_scroll_id = self.root.after(50, lambda: self._auto_scroll_during_drag(event) if self.is_dragging else None)
        elif canvas_y > canvas_height - scroll_zone and self.is_dragging:
            # Near bottom - scroll down (only if not already at bottom)
            first, last = self.file_list_canvas.yview()
            if last >= 1:
                if self.auto_scroll_id:
                    try:
                        self.root.after_cancel(self.auto_scroll_id)
                    except Exception:
                        pass
                    self.auto_scroll_id = None
                return
            self.file_list_canvas.yview_scroll(1, "units")
            # Schedule next scroll only if still dragging
            if self.auto_scroll_id and self.is_dragging:
                try:
                    self.root.after_cancel(self.auto_scroll_id)
                except Exception:
                    pass
            self.auto_scroll_id = self.root.after(50, lambda: self._auto_scroll_during_drag(event) if self.is_dragging else None)
        else:
            # In middle zone - cancel any pending scroll
            if self.auto_scroll_id:
                try:
                    self.root.after_cancel(self.auto_scroll_id)
                except Exception:
                    pass
                self.auto_scroll_id = None
    
    def _on_preview_toggle(self):
        """Hide preview when checkbox is unchecked"""
        if not self.preview_enabled.get():
            if self.preview_after_id:
                self.root.after_cancel(self.preview_after_id)
                self.preview_after_id = None
            self.pending_preview_index = None
            self.hide_preview()
        self._save_settings()
    
    def on_row_hover(self, event, index: int):
        """Handle row hover to show preview and update status bar"""
        # Update status bar with full path
        if 0 <= index < len(self.pdf_files):
            file_entry = self.pdf_files[index]
            file_path = self.get_file_path(file_entry)
            self.status_label.config(text=file_path)
        
        # Only show preview if enabled
        if not self.preview_enabled.get():
            return
        
        if 0 <= index < len(self.pdf_files):
            file_entry = self.pdf_files[index]
            file_path = self.get_file_path(file_entry)
            
            if self.preview_file_index != index:
                self._schedule_preview(index, file_path, event.x_root, event.y_root)
    
    def on_row_leave(self, event):
        """Hide preview and clear status bar when mouse leaves row"""
        self.status_label.config(text="")
        if self.preview_after_id:
            self.root.after_cancel(self.preview_after_id)
            self.preview_after_id = None
        self.pending_preview_index = None
        self.hide_preview()
    
    def _schedule_preview(self, index: int, file_path: str, x_root: int, y_root: int):
        """Schedule the preview popup with a short delay"""
        if self.preview_after_id:
            self.root.after_cancel(self.preview_after_id)
            self.preview_after_id = None
        
        self.pending_preview_index = index
        
        def _show_if_still_hovered():
            self.preview_after_id = None
            if not self.preview_enabled.get():
                return
            if self.pending_preview_index != index:
                return
            self.show_preview(index, x_root, y_root, file_path)
        
        self.preview_after_id = self.root.after(self.preview_delay_ms, _show_if_still_hovered)
    
    def on_row_double_click(self, event, index: int):
        """Open selected PDF file with system default viewer on double-click"""
        if 0 <= index < len(self.pdf_files):
            file_entry = self.pdf_files[index]
            file_path = self.get_file_path(file_entry)
            try:
                # Use os.startfile on Windows to open with default PDF viewer
                if sys.platform.startswith('win'):
                    os.startfile(file_path)
                elif sys.platform == 'darwin':  # macOS
                    os.system(f'open "{file_path}"')
                else:  # Linux
                    os.system(f'xdg-open "{file_path}"')
            except Exception as e:
                self.show_error_dialog("Error", f"Could not open file: {e}")
    
    def _update_row_visuals(self):
        """Update visual highlighting of selected rows - only updates rows that changed"""
        self.updating_visuals = True  # Prevent configure events
        
        rows = self.file_list_frame.winfo_children()
        
        # Use a noticeable but pleasant selection color
        for i, row in enumerate(rows):
            if hasattr(row, '_is_selected') and row._is_selected:
                bg_color = "#D0E8FF"  # Light blue - clearly visible
            else:
                bg_color = "white"
            
            # Only update if color actually changed
            if self.row_visual_state.get(i) != bg_color:
                row.config(bg=bg_color)
                # Update child labels
                for child in row.winfo_children():
                    if isinstance(child, (tk.Label, tk.Entry)):
                        try:
                            child.config(bg=bg_color)
                        except tk.TclError:
                            pass
                self.row_visual_state[i] = bg_color
        
        # Keep focus stable on canvas to prevent focus-change flicker
        self.file_list_canvas.focus_set()
        self.updating_visuals = False
        
        # Update button states based on selection
        self._update_button_states()
    
    def _update_button_states(self):
        """Enable/disable buttons based on current selection state"""
        # Check if any files are selected
        has_selection = False
        rows = self.file_list_frame.winfo_children()
        for row in rows:
            if hasattr(row, '_is_selected') and row._is_selected:
                has_selection = True
                break
        
        # Enable/disable Remove Selected button
        if has_selection:
            self.remove_button.config(state=tk.NORMAL)
        else:
            self.remove_button.config(state=tk.DISABLED)
        
        # Enable/disable Combine PDFs button (needs at least 2 files)
        if len(self.pdf_files) >= 2:
            self.combine_button.config(state=tk.NORMAL)
        else:
            self.combine_button.config(state=tk.DISABLED)
        
        
        # Enable/disable Clear All button (needs at least 1 file)
        if len(self.pdf_files) >= 1:
            self.clear_button.config(state=tk.NORMAL)
        else:
            self.clear_button.config(state=tk.DISABLED)
    
    def show_preview(self, index: int, x_root: int, y_root: int, file_path: str):
        """Show a preview popup with PDF or image thumbnail"""
        self.hide_preview()
        
        try:
            # Create preview window
            self.preview_window = tk.Toplevel(self.root)
            self.preview_window.wm_overrideredirect(True)
            self.preview_window.wm_attributes("-topmost", True)
            
            # Create main frame
            main_frame = tk.Frame(self.preview_window, bg="white", relief=tk.SOLID, borderwidth=1)
            main_frame.pack(padx=5, pady=5)
            
            # Get file info
            file_stat = os.stat(file_path)
            size_bytes = file_stat.st_size
            
            # Format file size
            if size_bytes < 1024:
                size_str = f"{size_bytes} B"
            elif size_bytes < 1024 * 1024:
                size_str = f"{size_bytes / 1024:.1f} KB"
            else:
                size_str = f"{size_bytes / (1024 * 1024):.1f} MB"
            
            # Check if file is an image or PDF
            is_image = file_path.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.tif'))
            
            if is_image:
                # Handle image file
                try:
                    img = Image.open(file_path)
                    # Convert RGBA to RGB if needed for preview
                    if img.mode == 'RGBA':
                        background = Image.new('RGB', img.size, (255, 255, 255))
                        background.paste(img, mask=img.split()[-1])
                        img = background
                    elif img.mode not in ('RGB', 'L'):
                        img = img.convert('RGB')
                    # Resize to fit preview window (max 180x240)
                    img.thumbnail((180, 240), Image.Resampling.LANCZOS)
                except Exception as e:
                    # Fallback if image loading fails
                    img = Image.new('RGB', (180, 240), color='#F0F0F0')
            else:
                # Handle PDF file
                try:
                    # Try to get page count for PDF
                    with open(file_path, 'rb') as pdf_file:
                        pdf_reader = PyPDF2.PdfReader(pdf_file)
                        page_count = len(pdf_reader.pages)
                    
                    # Convert first page to image using PyMuPDF
                    pdf_document = fitz.open(file_path)
                    if len(pdf_document) > 0:
                        # Render first page at 200 DPI
                        page = pdf_document[0]
                        pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
                        img_data = pix.tobytes("ppm")
                        img = Image.open(io.BytesIO(img_data))
                        # Resize to fit preview window (max 180x240)
                        img.thumbnail((180, 240), Image.Resampling.LANCZOS)
                    else:
                        img = Image.new('RGB', (180, 240), color='#F0F0F0')
                    pdf_document.close()
                except Exception as e:
                    # Fallback if PyMuPDF fails
                    img = Image.new('RGB', (180, 240), color='#F0F0F0')
            
            # Display the image using PhotoImage
            self.preview_photo = ImageTk.PhotoImage(img)
            img_label = tk.Label(main_frame, image=self.preview_photo, bg="white")
            img_label.pack(padx=5, pady=5)
            
            # Add filename only
            filename = Path(file_path).name
            filename_label = tk.Label(
                main_frame,
                text=filename,
                font=("Arial", 8),
                bg="white",
                fg="#000000",
                justify=tk.LEFT,
                wraplength=180
            )
            filename_label.pack(padx=5, pady=3)
            
            # Position near mouse
            x = x_root + 15
            y = y_root + 15
            self.preview_window.geometry(f"+{x}+{y}")
            
            self.preview_file_index = index
            
        except Exception as e:
            pass
    
    def hide_preview(self):
        """Hide the preview popup"""
        if self.preview_window:
            self.preview_window.destroy()
            self.preview_window = None
            self.preview_file_index = None
    
    def on_order_changed(self):
        """Deprecated: replaced by explicit sort controls."""
        pass

    def on_sort_clicked(self, key: str):
        """Handle sort header clicks. Clicking the same key toggles reverse; clicking a new key sets ascending."""
        if self.sort_key == key:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_key = key
            self.sort_reverse = False

        self.apply_sort()
        self.update_header_labels()

    def apply_sort(self):
        """Sort `self.pdf_files` according to current sort_key and sort_reverse, then refresh listbox."""
        try:
            if self.sort_key == 'name':
                self.pdf_files.sort(key=lambda x: Path(x['path']).name.lower(), reverse=self.sort_reverse)
            elif self.sort_key == 'size':
                def _size_key(entry):
                    try:
                        return os.path.getsize(entry['path'])
                    except Exception:
                        return -1
                self.pdf_files.sort(key=_size_key, reverse=self.sort_reverse)
            elif self.sort_key == 'date':
                def _date_key(entry):
                    try:
                        return os.path.getmtime(entry['path'])
                    except Exception:
                        return 0
                self.pdf_files.sort(key=_date_key, reverse=self.sort_reverse)
            # If sort_key is None, do nothing (preserve display order)

            # Update listbox (with numbering)
            self.refresh_listbox()
        except Exception:
            pass



    def update_header_labels(self):
        """Update header labels to show sort direction for the active key."""
        up = '▲'
        down = '▼'

        # Reset labels
        self.filename_hdr.config(text='Filename')
        self.size_hdr.config(text='Size')
        self.date_hdr.config(text='Date')

        if self.sort_key == 'name':
            arrow = down if self.sort_reverse else up
            self.filename_hdr.config(text=f'Filename {arrow}')
        elif self.sort_key == 'size':
            arrow = down if self.sort_reverse else up
            self.size_hdr.config(text=f'Size {arrow}')
        elif self.sort_key == 'date':
            arrow = down if self.sort_reverse else up
            self.date_hdr.config(text=f'Date {arrow}')
    
    
    def _create_text_instructions_tab(self, parent_frame):
        """Create instructions tab with markdown rendering and text formatting"""
        instructions_content_frame = tk.Frame(parent_frame)
        instructions_content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        scrollbar = tk.Scrollbar(instructions_content_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        instructions_text = tk.Text(
            instructions_content_frame,
            yscrollcommand=scrollbar.set,
            font=("Arial", 9),
            wrap=tk.WORD,
            bg="white",
            fg="black",
            relief=tk.FLAT,
            bd=0
        )
        instructions_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=instructions_text.yview)
        
        # Configure text tags for markdown rendering
        instructions_text.tag_configure("h1", font=("Arial", 12, "bold"), foreground="#0066CC", spacing1=6, spacing3=6)
        instructions_text.tag_configure("h2", font=("Arial", 11, "bold"), foreground="#0066CC", spacing1=4, spacing3=4)
        instructions_text.tag_configure("bold", font=("Arial", 9, "bold"))
        instructions_text.tag_configure("underline", font=("Arial", 9, "underline"))
        instructions_text.tag_configure("code", font=("Courier", 8), foreground="#666666", background="#F0F0F0")
        instructions_text.tag_configure("indent", lmargin1=20, lmargin2=20)
        
        # Helper to get resource path both when running normally and when frozen
        def resource_path(relative_path: str) -> str:
            if getattr(sys, "frozen", False):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(__file__)
            return os.path.join(base_path, relative_path)
        
        # Try to load markdown file
        try:
            md_file = resource_path("instructions.md")
            if os.path.exists(md_file):
                with open(md_file, 'r', encoding='utf-8') as f:
                    markdown_content = f.read()
                self._render_markdown(instructions_text, markdown_content)
            else:
                raise FileNotFoundError("instructions.md not found")
        except Exception as e:
            print(f"Warning: Could not load markdown instructions: {e}")
            # Fallback - show basic text
            instructions_text.insert(tk.END, "PDF Combiner Instructions\n\nSee the README or help documentation for usage instructions.")
        
        instructions_text.config(state=tk.DISABLED)  # Make read-only
    
    def _render_markdown(self, text_widget, markdown_content):
        """Render markdown content into a tk.Text widget with proper formatting"""
        lines = markdown_content.split('\n')
        
        for line in lines:
            if not line.strip():
                # Empty line
                text_widget.insert(tk.END, '\n')
            elif line.startswith('# '):
                # H1 header
                text_widget.insert(tk.END, line[2:] + '\n', 'h1')
            elif line.startswith('## '):
                # H2 header
                text_widget.insert(tk.END, line[3:] + '\n', 'h2')
            elif line.startswith('### '):
                # H3 header (subsection)
                text_widget.insert(tk.END, line[4:] + '\n', 'h2')
            elif line.startswith('- '):
                # Regular bullet point
                content = line[2:]
                text_widget.insert(tk.END, '• ')
                self._insert_markdown_line(text_widget, content, indent=False)
                text_widget.insert(tk.END, '\n')
            elif line.startswith('  - '):
                # Nested bullet point - apply indent tag to entire line
                content = line[4:]
                start_pos = text_widget.index(tk.END)
                text_widget.insert(tk.END, '    • ')
                self._insert_markdown_line(text_widget, content, indent=False)
                text_widget.insert(tk.END, '\n')
                # Apply indent tag to the entire bullet block
                end_pos = text_widget.index(tk.END)
                text_widget.tag_add('indent', start_pos, end_pos)
            else:
                # Regular text or continuation
                if line.startswith('  '):
                    # Indented text
                    start_pos = text_widget.index(tk.END)
                    self._insert_markdown_line(text_widget, line[2:], indent=False)
                    text_widget.insert(tk.END, '\n')
                    end_pos = text_widget.index(tk.END)
                    text_widget.tag_add('indent', start_pos, end_pos)
                else:
                    self._insert_markdown_line(text_widget, line, indent=False)
                    text_widget.insert(tk.END, '\n')
    
    def _insert_markdown_line(self, text_widget, line, indent=False):
        """Insert a single markdown line with bold, underline, and code formatting"""
        if indent:
            text_widget.insert(tk.END, '• ', 'indent')
        
        i = 0
        while i < len(line):
            # Look for underline markers (__text__)
            if line[i:i+2] == '__':
                # Find closing __
                close_idx = line.find('__', i + 2)
                if close_idx != -1:
                    text_widget.insert(tk.END, line[i+2:close_idx], 'underline')
                    i = close_idx + 2
                    continue
            
            # Look for bold markers
            if line[i:i+2] == '**':
                # Find closing **
                close_idx = line.find('**', i + 2)
                if close_idx != -1:
                    text_widget.insert(tk.END, line[i+2:close_idx], 'bold')
                    i = close_idx + 2
                    continue
            
            # Look for code markers
            if line[i] == '`':
                # Find closing `
                close_idx = line.find('`', i + 1)
                if close_idx != -1:
                    text_widget.insert(tk.END, line[i+1:close_idx], 'code')
                    i = close_idx + 1
                    continue
            
            # Regular text
            text_widget.insert(tk.END, line[i])
            i += 1
    
    def browse_output_location(self):
        """Open directory browser to select output location"""
        directory = filedialog.askdirectory(
            title="Select output location",
            initialdir=self.output_directory
        )
        
        if directory:
            self.output_directory = directory
            self.location_label.config(text=self.output_directory)
            self._save_settings()
    
    def _validate_output_filename(self, filename: str) -> tuple[bool, str, str]:
        """Validate output filename and return (is_valid, error_message, corrected_filename)"""
        import re
        
        # Store original for comparison
        original_filename = filename
        
        # Remove .pdf extension if present for validation
        has_pdf_ext = filename.lower().endswith('.pdf')
        if has_pdf_ext:
            base_filename = filename[:-4]
        else:
            base_filename = filename
        
        # Check if filename is empty
        if not base_filename.strip():
            return False, "Please enter a filename (cannot be empty or just whitespace).", original_filename
        
        # Check for invalid characters and replace them with underscores
        invalid_chars_pattern = r'[<>:"|?*\\]'
        corrected_base = re.sub(invalid_chars_pattern, '_', base_filename)
        had_invalid_chars = corrected_base != base_filename
        
        if had_invalid_chars:
            # Add back the .pdf extension if it was present
            corrected_filename = corrected_base + ('.pdf' if has_pdf_ext else '')
            return False, "Filename contains invalid characters: < > : \" | ? * \\\nThese will be replaced with underscores.", corrected_filename
        
        # Check filename length (Windows max is 255, minus .pdf extension)
        if len(corrected_base) > 240:
            return False, "Filename is too long (max 240 characters).", original_filename
        
        # Check for reserved Windows names
        reserved_names = {'con', 'prn', 'aux', 'nul', 'com1', 'com2', 'com3', 'com4', 'com5',
                         'com6', 'com7', 'com8', 'com9', 'lpt1', 'lpt2', 'lpt3', 'lpt4', 'lpt5',
                         'lpt6', 'lpt7', 'lpt8', 'lpt9'}
        if corrected_base.lower() in reserved_names:
            return False, f"'{corrected_base}' is a reserved filename. Please choose a different name.", original_filename
        
        return True, "", original_filename
    
    def combine_pdfs(self):
        """Combine selected PDF files"""
        if len(self.pdf_files) < 2:
            self.show_error_dialog("Error", "Please select at least 2 PDF files to combine.")
            return
        
        # Validate filename
        filename = self.output_filename.get().strip()
        
        # Validate the filename
        is_valid, error_message, corrected_filename = self._validate_output_filename(filename)
        if not is_valid:
            # For invalid characters, auto-correct and notify
            if "invalid characters" in error_message.lower():
                messagebox.showwarning("Filename Correction", error_message)
                self.output_filename.set(corrected_filename)
                filename = corrected_filename
            else:
                self.show_error_dialog("Invalid Filename", error_message)
                return
        
        # Ensure filename ends with .pdf
        if not filename.lower().endswith(".pdf"):
            filename += ".pdf"
        
        # Full output path
        output_file = str(Path(self.output_directory) / filename)

        # If output exists, ask user whether to overwrite
        if os.path.exists(output_file):
            if not self.show_overwrite_dialog(output_file):
                return

        # Get list of file entries to combine
        files_to_combine = self.pdf_files.copy()
        
        # Show summary before combining
        self.show_combine_summary(output_file, files_to_combine)
    
    def show_combine_summary(self, output_file, files_to_combine):
        """Show a summary of PDFs to combine before proceeding"""
        # Calculate total pages and file size
        original_pages = 0
        total_size_bytes = 0
        
        try:
            for file_entry in files_to_combine:
                file_path = self.get_file_path(file_entry)
                page_range = self.get_page_range(file_entry)
                
                # Check if this is an image file
                is_image = file_path.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.tif'))
                
                if is_image:
                    # Images convert to 1-page PDFs
                    total_file_pages = 1
                    page_indices = [0]  # Single page at index 0
                    original_pages += 1
                else:
                    # Count pages in PDF
                    with open(file_path, 'rb') as f:
                        pdf_reader = PyPDF2.PdfReader(f)
                        total_file_pages = len(pdf_reader.pages)
                        try:
                            page_indices = self._parse_page_range(page_range, total_file_pages)
                            original_pages += len(page_indices)
                        except ValueError as e:
                            error_msg = (
                                f"{Path(file_path).name} (Total pages: {total_file_pages})\n\n"
                                f"Error: {e}\n\n"
                                f"Valid formats:\n"
                                f"  • All pages: 'All' or leave blank\n"
                                f"  • Single page: '5'\n"
                                f"  • Range: '1-10'\n"
                                f"  • Multiple ranges: '1-3,5,7-9'"
                            )
                            self.show_error_dialog("Invalid Page Range", error_msg)
                            return
                
                # Get file size
                total_size_bytes += os.path.getsize(file_path)
        except Exception as e:
            self.show_error_dialog("Error", f"Could not read file information: {e}")
            return
        
        # Calculate blank pages if enabled (inserted between files)
        blank_pages = 0
        if self.insert_blank_pages.get() and len(files_to_combine) > 1:
            blank_pages = len(files_to_combine) - 1
        total_pages = original_pages + blank_pages

        # Format file size
        if total_size_bytes < 1024:
            size_str = f"{total_size_bytes} B"
        elif total_size_bytes < 1024 * 1024:
            size_str = f"{total_size_bytes / 1024:.1f} KB"
        else:
            size_str = f"{total_size_bytes / (1024 * 1024):.1f} MB"
        
        # Create summary window
        summary_window = tk.Toplevel(self.root)
        summary_window.title("Combine Summary")
        dialog_width, dialog_height = self._scale_geometry(550, 500)
        summary_window.geometry(f"{dialog_width}x{dialog_height}")
        summary_window.resizable(False, False)
        summary_window.transient(self.root)
        summary_window.grab_set()
        
        # Center the summary window on parent window
        summary_window.update_idletasks()
        parent_x = self.root.winfo_x()
        parent_y = self.root.winfo_y()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()
        center_x = parent_x + (parent_width - dialog_width) // 2
        top_y = parent_y
        summary_window.geometry(f"{dialog_width}x{dialog_height}+{center_x}+{top_y}")
        
        # Title
        title_label = tk.Label(
            summary_window,
            text="Combine Summary",
            font=("Arial", 12, "bold"),
            pady=10
        )
        title_label.pack()
        
        # Info frame with table layout
        info_frame = tk.Frame(summary_window)
        info_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        # Configure grid columns: left for labels, right for values
        info_frame.columnconfigure(0, weight=0)  # Labels column
        info_frame.columnconfigure(1, weight=1)  # Values column
        
        row = 0
        
        # Files count
        tk.Label(info_frame, text="Files to combine:", font=("Arial", 10, "bold"), fg="black", anchor="e", justify=tk.RIGHT).grid(row=row, column=0, sticky="ne", padx=(0, 15), pady=5)
        tk.Label(info_frame, text=f"{len(files_to_combine)} files", font=("Arial", 10, "bold"), fg="#0066CC", anchor="w").grid(row=row, column=1, sticky="nw", pady=5)
        row += 1
        
        # Total pages
        if blank_pages > 0:
            pages_text = f"{total_pages} pages ({original_pages} from PDFs + {blank_pages} breaker pages)"
        else:
            pages_text = f"{total_pages} pages"
        tk.Label(info_frame, text="Total pages:", font=("Arial", 10, "bold"), fg="black", anchor="e", justify=tk.RIGHT).grid(row=row, column=0, sticky="ne", padx=(0, 15), pady=5)
        tk.Label(info_frame, text=pages_text, font=("Arial", 10, "bold"), fg="#0066CC", anchor="w").grid(row=row, column=1, sticky="nw", pady=5)
        row += 1
        
        # Total size
        tk.Label(info_frame, text="Total size:", font=("Arial", 10, "bold"), fg="black", anchor="e", justify=tk.RIGHT).grid(row=row, column=0, sticky="ne", padx=(0, 15), pady=5)
        tk.Label(info_frame, text=size_str, font=("Arial", 10, "bold"), fg="#0066CC", anchor="w").grid(row=row, column=1, sticky="nw", pady=5)
        row += 1
        
        # Save path
        tk.Label(info_frame, text="Save to:", font=("Arial", 10, "bold"), fg="black", anchor="ne", justify=tk.RIGHT).grid(row=row, column=0, sticky="ne", padx=(0, 15), pady=5)
        tk.Label(info_frame, text=output_file, font=("Arial", 9, "bold"), fg="#0066CC", anchor="w", wraplength=300, justify=tk.LEFT).grid(row=row, column=1, sticky="nw", pady=5)
        row += 1
        
        # Compression level
        tk.Label(info_frame, text="Compression:", font=("Arial", 9, "bold"), fg="black", anchor="e", justify=tk.RIGHT).grid(row=row, column=0, sticky="ne", padx=(0, 15), pady=4)
        tk.Label(info_frame, text=self.compression_quality.get(), font=("Arial", 9, "bold"), fg="#0066CC", anchor="w").grid(row=row, column=1, sticky="nw", pady=4)
        row += 1
        
        # Bookmarks
        bookmarks_enabled = self.add_filename_bookmarks.get()
        tk.Label(info_frame, text="Bookmarks:", font=("Arial", 9, "bold"), fg="black", anchor="e", justify=tk.RIGHT).grid(row=row, column=0, sticky="ne", padx=(0, 15), pady=4)
        tk.Label(info_frame, text=f"{'Enabled' if bookmarks_enabled else 'Disabled'}", font=("Arial", 9, "bold"), fg="#006600" if bookmarks_enabled else "#CC0000", anchor="w").grid(row=row, column=1, sticky="nw", pady=4)
        row += 1
        
        # Table of Contents
        toc_enabled = self.insert_toc.get()
        tk.Label(info_frame, text="Table of Contents:", font=("Arial", 9, "bold"), fg="black", anchor="e", justify=tk.RIGHT).grid(row=row, column=0, sticky="ne", padx=(0, 15), pady=4)
        tk.Label(info_frame, text=f"{'Enabled' if toc_enabled else 'Disabled'}", font=("Arial", 9, "bold"), fg="#006600" if toc_enabled else "#CC0000", anchor="w").grid(row=row, column=1, sticky="nw", pady=4)
        row += 1
        
        # Watermark
        watermark_enabled = self.enable_watermark.get()
        if watermark_enabled:
            watermark_text = f"Enabled - '{self.watermark_text.get()}' ({self.watermark_position.get().lower()})"
            watermark_color = "#006600"
        else:
            watermark_text = "Disabled"
            watermark_color = "#CC0000"
        tk.Label(info_frame, text="Watermark:", font=("Arial", 9, "bold"), fg="black", anchor="ne", justify=tk.RIGHT).grid(row=row, column=0, sticky="ne", padx=(0, 15), pady=4)
        tk.Label(info_frame, text=watermark_text, font=("Arial", 9, "bold"), fg=watermark_color, anchor="w", wraplength=300, justify=tk.LEFT).grid(row=row, column=1, sticky="nw", pady=4)
        row += 1
        
        # Page Scaling
        scaling_enabled = self.enable_page_scaling.get()
        tk.Label(info_frame, text="Page Scaling:", font=("Arial", 9, "bold"), fg="black", anchor="e", justify=tk.RIGHT).grid(row=row, column=0, sticky="ne", padx=(0, 15), pady=4)
        tk.Label(info_frame, text=f"{'Enabled' if scaling_enabled else 'Disabled'}", font=("Arial", 9, "bold"), fg="#006600" if scaling_enabled else "#CC0000", anchor="w").grid(row=row, column=1, sticky="nw", pady=4)
        row += 1
        
        # Insert breaker pages
        breaker_enabled = self.insert_blank_pages.get()
        tk.Label(info_frame, text="Insert breaker pages:", font=("Arial", 9, "bold"), fg="black", anchor="e", justify=tk.RIGHT).grid(row=row, column=0, sticky="ne", padx=(0, 15), pady=4)
        tk.Label(info_frame, text=f"{'Enabled' if breaker_enabled else 'Disabled'}", font=("Arial", 9, "bold"), fg="#006600" if breaker_enabled else "#CC0000", anchor="w").grid(row=row, column=1, sticky="nw", pady=4)
        row += 1
        
        # Ignore blank pages
        ignore_blank_enabled = self.delete_blank_pages.get()
        tk.Label(info_frame, text="Ignore blank pages:", font=("Arial", 9, "bold"), fg="black", anchor="e", justify=tk.RIGHT).grid(row=row, column=0, sticky="ne", padx=(0, 15), pady=4)
        tk.Label(info_frame, text=f"{'Enabled' if ignore_blank_enabled else 'Disabled'}", font=("Arial", 9, "bold"), fg="#006600" if ignore_blank_enabled else "#CC0000", anchor="w").grid(row=row, column=1, sticky="nw", pady=4)
        row += 1
        
        # Add PDF metadata
        metadata_enabled = self.enable_metadata.get()
        tk.Label(info_frame, text="Add PDF metadata:", font=("Arial", 9, "bold"), fg="black", anchor="e", justify=tk.RIGHT).grid(row=row, column=0, sticky="ne", padx=(0, 15), pady=4)
        tk.Label(info_frame, text=f"{'Enabled' if metadata_enabled else 'Disabled'}", font=("Arial", 9, "bold"), fg="#006600" if metadata_enabled else "#CC0000", anchor="w").grid(row=row, column=1, sticky="nw", pady=4)
        row += 1
        
        # Button frame
        button_frame = tk.Frame(summary_window)
        button_frame.pack(pady=10)
        
        # Proceed button
        proceed_button = tk.Button(
            button_frame,
            text="Proceed",
            command=lambda: (
                summary_window.destroy(),
                self.show_progress_dialog(output_file, files_to_combine)
            ),
            width=12,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10)
        )
        proceed_button.grid(row=0, column=0, padx=5)
        
        # Cancel button
        cancel_button = tk.Button(
            button_frame,
            text="Cancel",
            command=summary_window.destroy,
            width=12,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10)
        )
        cancel_button.grid(row=0, column=1, padx=5)
    
    def show_progress_dialog(self, output_file, files_to_combine):
        """Show a progress dialog while combining PDFs"""
        # Create progress window
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Combining PDFs")
        dialog_width, dialog_height = self._scale_geometry(400, 190)
        progress_window.geometry(f"{dialog_width}x{dialog_height}")
        progress_window.resizable(False, False)
        progress_window.transient(self.root)
        progress_window.grab_set()
        
        # Center the progress window on parent window
        progress_window.update_idletasks()
        parent_x = self.root.winfo_x()
        parent_y = self.root.winfo_y()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()
        center_x = parent_x + (parent_width - dialog_width) // 2
        center_y = parent_y + (parent_height - dialog_height) // 2
        progress_window.geometry(f"{dialog_width}x{dialog_height}+{center_x}+{center_y}")
        
        # Progress label
        progress_label = tk.Label(
            progress_window,
            text="Preparing to combine PDFs...",
            font=("Arial", 10),
            pady=10
        )
        progress_label.pack()
        
        # Progress bar
        progress_bar = ttk.Progressbar(
            progress_window,
            mode='determinate',
            length=350,
            maximum=len(files_to_combine) + 1
        )
        progress_bar.pack(pady=10)
        
        # File counter label
        counter_label = tk.Label(
            progress_window,
            text=f"0 of {len(files_to_combine)} files processed",
            font=("Arial", 9),
            fg="#666666"
        )
        counter_label.pack(pady=5)
        
        # Cancel flag
        cancel_flag = {'cancelled': False}
        
        # Cancel button
        def on_cancel():
            cancel_flag['cancelled'] = True
            cancel_button.config(state='disabled', text="Cancelling...")
        
        cancel_button = tk.Button(
            progress_window,
            text="Cancel",
            command=on_cancel,
            width=15,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 9)
        )
        cancel_button.pack(pady=10)
        
        # Run combine operation in thread
        def combine_thread():
            pdf_writer = None
            temp_pdf_files = []  # Track temporary PDF files created from images for cleanup
            try:
                # Create PDF writer object
                pdf_writer = PyPDF2.PdfWriter()
                
                # Track current page number for bookmarks
                current_page_num = 0
                
                # Track file TOC entries (filename, starting page) for TOC generation
                file_toc_entries = []
                
                # First pass: determine max dimensions if scaling is enabled
                max_width = 0
                max_height = 0
                if self.enable_page_scaling.get():
                    for file_entry in files_to_combine:
                        file_path = self.get_file_path(file_entry)
                        rotation = self.get_rotation(file_entry)
                        page_range = self.get_page_range(file_entry)
                        
                        # Check if this is an image file
                        is_image = file_path.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.tif'))
                        
                        if is_image:
                            # Get image dimensions and convert to PDF points
                            try:
                                img = Image.open(file_path)
                                # Must use same DPI as _image_to_pdf() method
                                dpi = 96
                                width = (float(img.width) / dpi) * 72  # Convert pixels to points
                                height = (float(img.height) / dpi) * 72
                                if rotation in [90, 270]:
                                    width, height = height, width
                                max_width = max(max_width, width)
                                max_height = max(max_height, height)
                            except Exception:
                                # If we can't read image dimensions, use default letter size
                                max_width = max(max_width, 612.0)
                                max_height = max(max_height, 792.0)
                        else:
                            # Get PDF dimensions
                            with open(file_path, 'rb') as pdf_file:
                                pdf_reader = PyPDF2.PdfReader(pdf_file)
                                total_file_pages = len(pdf_reader.pages)
                                try:
                                    page_indices = self._parse_page_range(page_range, total_file_pages)
                                except ValueError:
                                    continue
                                
                                # Determine if a specific page range was selected (not "All")
                                has_explicit_range = (page_range and 
                                                    page_range.strip().lower() not in ["all", ""])
                                
                                # Check dimensions of each page
                                for page_index in page_indices:
                                    page = pdf_reader.pages[page_index]
                                    # Skip blank pages if that option is enabled AND no explicit range was selected
                                    if self.delete_blank_pages.get() and not has_explicit_range and self._is_page_blank(page):
                                        continue
                                    
                                    box = page.mediabox
                                    width = float(box.width)
                                    height = float(box.height)
                                    if rotation in [90, 270]:
                                        width, height = height, width
                                    max_width = max(max_width, width)
                                    max_height = max(max_height, height)
                
                # Second pass: process and add pages
                for i, file_entry in enumerate(files_to_combine):
                    # Check if cancelled
                    if cancel_flag['cancelled']:
                        self.root.after(0, lambda: (
                            progress_window.destroy(),
                            self.show_info_dialog("Cancelled", "PDF combining operation was cancelled.")
                        ))
                        return
                    
                    file_path = self.get_file_path(file_entry)
                    rotation = self.get_rotation(file_entry)
                    page_range = self.get_page_range(file_entry)
                    reverse = self.get_reverse(file_entry)
                    
                    # Update progress
                    self.root.after(0, lambda idx=i, f=file_path: (
                        progress_label.config(text=f"Processing: {Path(f).name}"),
                        progress_bar.config(value=idx),
                        counter_label.config(text=f"{idx} of {len(files_to_combine)} files processed")
                    ))
                    
                    # Convert image to PDF if necessary
                    pdf_path = file_path
                    is_image = file_path.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.tif'))
                    
                    if is_image:
                        # Convert image to PDF
                        try:
                            pdf_path = self._image_to_pdf(file_path)
                            temp_pdf_files.append(pdf_path)  # Track for cleanup
                        except Exception as e:
                            # Cleanup any files created so far before showing error
                            for temp_file in temp_pdf_files:
                                try:
                                    os.remove(temp_file)
                                except:
                                    pass
                            self.root.after(0, lambda: (
                                progress_window.destroy(),
                                self.show_error_dialog("Image Conversion Error", f"Failed to convert image {Path(file_path).name}:\n{str(e)}")
                            ))
                            return
                    
                    # Read and process the PDF
                    with open(pdf_path, 'rb') as pdf_file:
                        pdf_reader = PyPDF2.PdfReader(pdf_file)
                        total_file_pages = len(pdf_reader.pages)
                        
                        # Insert first breaker page if enabled
                        if i == 0 and self.insert_blank_pages.get():
                            breaker_width = float(pdf_reader.pages[0].mediabox.width) if len(pdf_reader.pages) > 0 else 612
                            breaker_height = float(pdf_reader.pages[0].mediabox.height) if len(pdf_reader.pages) > 0 else 792
                            # Swap dimensions if file has 90 or 270 degree rotation
                            if rotation in [90, 270]:
                                breaker_width, breaker_height = breaker_height, breaker_width
                            first_filename = Path(file_path).name
                            breaker_page = self._create_page_with_filename(first_filename, breaker_width, breaker_height)
                            pdf_writer.add_page(breaker_page)
                            current_page_num += 1
                        try:
                            page_indices = self._parse_page_range(page_range, total_file_pages)
                        except ValueError as e:
                            error_msg = (
                                f"{Path(file_path).name} (Total pages: {total_file_pages})\n\n"
                                f"Error: {e}\n\n"
                                f"Valid formats:\n"
                                f"  • All pages: 'All' or leave blank\n"
                                f"  • Single page: '5'\n"
                                f"  • Range: '1-10'\n"
                                f"  • Multiple ranges: '1-3,5,7-9'"
                            )
                            self.root.after(0, lambda: (
                                progress_window.destroy(),
                                self.show_error_dialog("Invalid Page Range", error_msg)
                            ))
                            return
                        
                        # Reverse page indices if requested
                        if reverse:
                            page_indices = list(reversed(page_indices))
                        
                        # Determine if a specific page range was selected (not "All")
                        # Empty, None, or "all" means user selected all pages
                        has_explicit_range = (page_range and 
                                            page_range.strip().lower() not in ["all", ""])
                        
                        # Track file starting page for TOC if enabled (before adding breaker page)
                        file_start_page = current_page_num
                        
                        # Add bookmark at the start of this file's pages
                        parent_bookmark = None
                        if self.add_filename_bookmarks.get() and len(page_indices) > 0:
                            bookmark_title = Path(file_path).stem
                            parent_bookmark = pdf_writer.add_outline_item(bookmark_title, current_page_num)
                        
                        # Process each page
                        for page_index in page_indices:
                            page = pdf_reader.pages[page_index]
                            
                            # Skip blank pages if enabled AND no explicit range was selected
                            # This respects explicit page ranges while still filtering blanks for "All"
                            if self.delete_blank_pages.get() and not has_explicit_range:
                                if self._is_page_blank(page):
                                    continue
                            
                            # Apply rotation if specified
                            if rotation != 0:
                                page.rotate(rotation)
                            
                            # Scale page to uniform size if enabled
                            if self.enable_page_scaling.get() and max_width > 0 and max_height > 0:
                                self._scale_page(page, max_width, max_height)
                            
                            # Add watermark if enabled
                            if self.enable_watermark.get() and self.watermark_text.get().strip():
                                self._add_watermark(page, self.watermark_text.get(), self.watermark_opacity.get(), self.watermark_font_size.get(), self.watermark_rotation.get(), self.watermark_position.get())
                            
                            pdf_writer.add_page(page)
                            current_page_num += 1
                        
                        # Insert breaker page between files if enabled (but not after the last file)
                        if self.insert_blank_pages.get() and i < len(files_to_combine) - 1:
                            # Get the next file's information
                            next_file_entry = files_to_combine[i + 1]
                            next_file_path = self.get_file_path(next_file_entry)
                            next_filename = Path(next_file_path).name
                            next_rotation = self.get_rotation(next_file_entry)
                            
                            # Get breaker page dimensions
                            blank_width = 612.0
                            blank_height = 792.0
                            
                            # If not using uniform size, match the next file's dimensions
                            if not self.breaker_pages_uniform_size.get():
                                try:
                                    next_is_image = next_file_path.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.tif'))
                                    
                                    if next_is_image:
                                        # Get image dimensions and convert to PDF points
                                        next_img = Image.open(next_file_path)
                                        # Must use same DPI as _image_to_pdf() method
                                        dpi = 96
                                        blank_width = (float(next_img.width) / dpi) * 72  # Convert pixels to points
                                        blank_height = (float(next_img.height) / dpi) * 72
                                        # Swap dimensions if next file has 90 or 270 degree rotation
                                        if next_rotation in [90, 270]:
                                            blank_width, blank_height = blank_height, blank_width
                                    else:
                                        # Get PDF dimensions
                                        with open(next_file_path, 'rb') as next_pdf_file:
                                            next_pdf_reader = PyPDF2.PdfReader(next_pdf_file)
                                            if len(next_pdf_reader.pages) > 0:
                                                next_page = next_pdf_reader.pages[0]
                                                blank_width = float(next_page.mediabox.width)
                                                blank_height = float(next_page.mediabox.height)
                                                # Swap dimensions if next file has 90 or 270 degree rotation
                                                if next_rotation in [90, 270]:
                                                    blank_width, blank_height = blank_height, blank_width
                                except:
                                    # Fallback to default letter size if we can't read the next file
                                    pass
                            
                            # Create blank page with filename text and add it
                            blank_page = self._create_page_with_filename(next_filename, blank_width, blank_height)
                            pdf_writer.add_page(blank_page)
                            current_page_num += 1
                        
                        # Add file to TOC entries (after processing this file)
                        if self.insert_toc.get() and len(page_indices) > 0:
                            file_toc_entries.append({
                                'filename': Path(file_path).name,
                                'page': file_start_page
                            })
                
                # Check if cancelled before writing
                if cancel_flag['cancelled']:
                    self.root.after(0, lambda: (
                        progress_window.destroy(),
                        self.show_info_dialog("Cancelled", "PDF combining operation was cancelled.")
                    ))
                    return
                
                # Add metadata if enabled
                if self.enable_metadata.get() and (self.pdf_title.get() or self.pdf_author.get() or self.pdf_subject.get() or self.pdf_keywords.get()):
                    metadata = {}
                    if self.pdf_title.get():
                        metadata['/Title'] = self.pdf_title.get()
                    if self.pdf_author.get():
                        metadata['/Author'] = self.pdf_author.get()
                    if self.pdf_subject.get():
                        metadata['/Subject'] = self.pdf_subject.get()
                    if self.pdf_keywords.get():
                        metadata['/Keywords'] = self.pdf_keywords.get()
                    pdf_writer.add_metadata(metadata)
                
                # Update for writing phase
                self.root.after(0, lambda: (
                    progress_label.config(text="Writing combined PDF..."),
                    progress_bar.config(value=len(files_to_combine)),
                    counter_label.config(text=f"{len(files_to_combine)} of {len(files_to_combine)} files processed"),
                    cancel_button.config(state='disabled')
                ))
                
                # Write combined PDF with compression settings
                with open(output_file, 'wb') as out_file:
                    # Apply compression if enabled
                    compression_level = self.compression_quality.get()
                    if compression_level != "None":
                        for page in pdf_writer.pages:
                            self._compress_page(page, compression_level)
                    pdf_writer.write(out_file)
                
                # Insert TOC if enabled
                if self.insert_toc.get() and len(file_toc_entries) > 0:
                    self.root.after(0, lambda: progress_label.config(text="Inserting Table of Contents..."))
                    self._insert_toc_page(output_file, file_toc_entries)
                elif self.insert_toc.get():
                    pass  # TOC checkbox enabled but no entries to add
                
                # Remember the output file
                self.last_output_file = output_file
                
                # Close progress window and show success
                self.root.after(0, lambda: (
                    progress_window.destroy(),
                    self.show_success_dialog(output_file)
                ))
                
            except FileNotFoundError as e:
                error_msg = f"File not found: {e}"
                self.root.after(0, lambda: (
                    progress_window.destroy(),
                    self.show_error_dialog("Error", error_msg)
                ))
            except Exception as e:
                error_msg = f"An error occurred: {str(e)}"
                self.root.after(0, lambda: (
                    progress_window.destroy(),
                    self.show_error_dialog("Error", error_msg)
                ))
            except Exception as e:
                error_msg = f"An error occurred: {str(e)}"
                self.root.after(0, lambda: (
                    progress_window.destroy(),
                    self.show_error_dialog("Error", error_msg)
                ))
            finally:
                # Clean up temporary PDF files created from images
                for temp_file in temp_pdf_files:
                    try:
                        os.remove(temp_file)
                    except:
                        pass  # Silently ignore cleanup errors
        
        # Start the thread
        thread = threading.Thread(target=combine_thread, daemon=True)
        thread.start()
    
    def show_success_dialog(self, output_file):
        """Show success dialog centered on parent and ask to open the file"""
        # Create success window
        success_window = tk.Toplevel(self.root)
        success_window.title("Success")
        dialog_width, dialog_height = self._scale_geometry(500, 200)
        success_window.geometry(f"{dialog_width}x{dialog_height}")
        success_window.resizable(False, False)
        success_window.transient(self.root)
        success_window.grab_set()
        
        # Center the success window on parent window
        success_window.update_idletasks()
        parent_x = self.root.winfo_x()
        parent_y = self.root.winfo_y()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()
        center_x = parent_x + (parent_width - dialog_width) // 2
        center_y = parent_y + (parent_height - dialog_height) // 2
        success_window.geometry(f"{dialog_width}x{dialog_height}+{center_x}+{center_y}")
        
        # Title
        title_label = tk.Label(
            success_window,
            text="✓ PDFs Combined Successfully!",
            font=("Arial", 12, "bold"),
            fg="#006600",
            pady=10
        )
        title_label.pack()
        
        # Info frame
        info_frame = tk.Frame(success_window)
        info_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        # Message
        message_label = tk.Label(
            info_frame,
            text="Your PDFs have been combined successfully.",
            font=("Arial", 10),
            fg="black",
            anchor="w",
            justify=tk.LEFT
        )
        message_label.pack(fill=tk.X, pady=5)
        
        # File path
        path_label = tk.Label(
            info_frame,
            text="Combined PDF saved to:",
            font=("Arial", 9, "bold"),
            fg="black",
            anchor="w"
        )
        path_label.pack(fill=tk.X, pady=(10, 2))
        
        path_value = tk.Label(
            info_frame,
            text=output_file,
            font=("Arial", 9),
            fg="#0066CC",
            anchor="w",
            wraplength=450,
            justify=tk.LEFT
        )
        path_value.pack(fill=tk.X, padx=10)
        
        # Button frame
        button_frame = tk.Frame(success_window)
        button_frame.pack(pady=10)
        
        # Open button
        open_button = tk.Button(
            button_frame,
            text="Open PDF",
            command=lambda: (
                success_window.destroy(),
                self._open_success_file(output_file)
            ),
            width=14,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10)
        )
        open_button.grid(row=0, column=0, padx=5)
        
        # Close button
        close_button = tk.Button(
            button_frame,
            text="Close",
            command=success_window.destroy,
            width=14,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10)
        )
        close_button.grid(row=0, column=1, padx=5)
    
    def _open_success_file(self, output_file):
        """Helper to open the combined PDF file"""
        try:
            if os.name == 'nt':
                os.startfile(output_file)
            else:
                webbrowser.open_new(output_file)
        except Exception as e:
            self.show_error_dialog("Error", f"Could not open file: {e}")
    
    def show_overwrite_dialog(self, output_file):
        """Show overwrite dialog centered on parent window and return True/False"""
        # Use a list to capture the result from the nested window
        result = [False]
        
        # Create overwrite window
        overwrite_window = tk.Toplevel(self.root)
        overwrite_window.title("Overwrite File")
        dialog_width, dialog_height = self._scale_geometry(480, 220)
        overwrite_window.geometry(f"{dialog_width}x{dialog_height}")
        overwrite_window.resizable(False, False)
        overwrite_window.transient(self.root)
        overwrite_window.grab_set()
        
        # Center the overwrite window on parent window
        overwrite_window.update_idletasks()
        parent_x = self.root.winfo_x()
        parent_y = self.root.winfo_y()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()
        center_x = parent_x + (parent_width - dialog_width) // 2
        center_y = parent_y + (parent_height - dialog_height) // 2
        overwrite_window.geometry(f"{dialog_width}x{dialog_height}+{center_x}+{center_y}")
        
        # Title
        title_label = tk.Label(
            overwrite_window,
            text="File Already Exists",
            font=("Arial", 12, "bold"),
            fg="black",
            pady=10
        )
        title_label.pack()
        
        # Info frame
        info_frame = tk.Frame(overwrite_window)
        info_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        # Message
        message_label = tk.Label(
            info_frame,
            text="The file already exists:",
            font=("Arial", 10),
            fg="black",
            anchor="w",
            justify=tk.LEFT
        )
        message_label.pack(fill=tk.X, pady=(0, 5))
        
        # File path
        path_value = tk.Label(
            info_frame,
            text=output_file,
            font=("Arial", 9),
            fg="#0066CC",
            anchor="w",
            wraplength=430,
            justify=tk.LEFT
        )
        path_value.pack(fill=tk.X, padx=10, pady=5)
        
        # Question
        question_label = tk.Label(
            info_frame,
            text="Do you want to overwrite it?",
            font=("Arial", 10),
            fg="black",
            anchor="w"
        )
        question_label.pack(fill=tk.X, pady=(10, 0))
        
        # Button frame
        button_frame = tk.Frame(overwrite_window)
        button_frame.pack(pady=10)
        
        # Overwrite button
        overwrite_button = tk.Button(
            button_frame,
            text="Overwrite",
            command=lambda: (
                result.__setitem__(0, True),
                overwrite_window.destroy()
            ),
            width=14,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10)
        )
        overwrite_button.grid(row=0, column=0, padx=5)
        
        # Cancel button
        cancel_button = tk.Button(
            button_frame,
            text="Cancel",
            command=overwrite_window.destroy,
            width=14,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10)
        )
        cancel_button.grid(row=0, column=1, padx=5)
        
        # Wait for window to close and return result
        self.root.wait_window(overwrite_window)
        return result[0]
    
    def show_info_dialog(self, title, message):
        """Show an info/success dialog centered on parent window"""
        info_window = tk.Toplevel(self.root)
        info_window.title(title)
        dialog_width, dialog_height = self._scale_geometry(450, 170)
        info_window.geometry(f"{dialog_width}x{dialog_height}")
        info_window.resizable(False, False)
        info_window.transient(self.root)
        info_window.grab_set()
        
        # Center the info window on parent window
        info_window.update_idletasks()
        parent_x = self.root.winfo_x()
        parent_y = self.root.winfo_y()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()
        center_x = parent_x + (parent_width - dialog_width) // 2
        center_y = parent_y + (parent_height - dialog_height) // 2
        info_window.geometry(f"{dialog_width}x{dialog_height}+{center_x}+{center_y}")
        
        # Title
        title_label = tk.Label(
            info_window,
            text=title,
            font=("Arial", 12, "bold"),
            fg="#006600",
            pady=8
        )
        title_label.pack()
        
        # Message
        message_label = tk.Label(
            info_window,
            text=message,
            font=("Arial", 10),
            fg="black",
            anchor="w",
            justify=tk.LEFT,
            wraplength=400
        )
        message_label.pack(pady=5, padx=20, fill=tk.BOTH, expand=True)
        
        # OK button
        button_frame = tk.Frame(info_window)
        button_frame.pack(pady=(5, 10))
        
        ok_button = tk.Button(
            button_frame,
            text="OK",
            command=info_window.destroy,
            width=14,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10)
        )
        ok_button.pack()
        
        self.root.wait_window(info_window)
    
    def show_error_dialog(self, title, message):
        """Show an error dialog centered on parent window"""
        error_window = tk.Toplevel(self.root)
        error_window.title(title)
        dialog_width, dialog_height = self._scale_geometry(450, 180)
        error_window.geometry(f"{dialog_width}x{dialog_height}")
        error_window.resizable(False, False)
        error_window.transient(self.root)
        error_window.grab_set()
        
        # Center the error window on parent window
        error_window.update_idletasks()
        parent_x = self.root.winfo_x()
        parent_y = self.root.winfo_y()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()
        center_x = parent_x + (parent_width - dialog_width) // 2
        center_y = parent_y + (parent_height - dialog_height) // 2
        error_window.geometry(f"{dialog_width}x{dialog_height}+{center_x}+{center_y}")
        
        # Title
        title_label = tk.Label(
            error_window,
            text=title,
            font=("Arial", 12, "bold"),
            fg="#CC0000",
            pady=10
        )
        title_label.pack()
        
        # Message
        message_label = tk.Label(
            error_window,
            text=message,
            font=("Arial", 10),
            fg="black",
            anchor="w",
            justify=tk.LEFT,
            wraplength=400
        )
        message_label.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        # OK button
        button_frame = tk.Frame(error_window)
        button_frame.pack(pady=10)
        
        ok_button = tk.Button(
            button_frame,
            text="OK",
            command=error_window.destroy,
            width=14,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10)
        )
        ok_button.pack()
        
        self.root.wait_window(error_window)
    
    def show_merge_lists_dialog(self, valid_entries_count):
        """Show merge lists dialog centered on parent window. Returns True (append), False (replace), or None (cancel)"""
        result = [None]  # Use list to capture result from nested window
        
        # Create merge dialog window
        merge_window = tk.Toplevel(self.root)
        merge_window.title("Merge Lists?")
        dialog_width, dialog_height = self._scale_geometry(480, 220)
        merge_window.geometry(f"{dialog_width}x{dialog_height}")
        merge_window.resizable(False, False)
        merge_window.transient(self.root)
        merge_window.grab_set()
        
        # Center the merge window on parent window
        merge_window.update_idletasks()
        parent_x = self.root.winfo_x()
        parent_y = self.root.winfo_y()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()
        center_x = parent_x + (parent_width - dialog_width) // 2
        center_y = parent_y + (parent_height - dialog_height) // 2
        merge_window.geometry(f"{dialog_width}x{dialog_height}+{center_x}+{center_y}")
        
        # Title
        title_label = tk.Label(
            merge_window,
            text="Merge PDF Lists?",
            font=("Arial", 12, "bold"),
            fg="black",
            pady=10
        )
        title_label.pack()
        
        # Info frame
        info_frame = tk.Frame(merge_window)
        info_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        # Message
        message_label = tk.Label(
            info_frame,
            text=f"Found {valid_entries_count} valid PDFs in the saved list.",
            font=("Arial", 10),
            fg="black",
            anchor="w",
            justify=tk.LEFT
        )
        message_label.pack(fill=tk.X, pady=(0, 10))
        
        # Options
        options_label = tk.Label(
            info_frame,
            text="What would you like to do?",
            font=("Arial", 9),
            fg="#666666",
            anchor="w"
        )
        options_label.pack(fill=tk.X, pady=(0, 5))
        
        option1_label = tk.Label(
            info_frame,
            text="• Append: Add to current list",
            font=("Arial", 9),
            fg="black",
            anchor="w"
        )
        option1_label.pack(fill=tk.X, padx=10)
        
        option2_label = tk.Label(
            info_frame,
            text="• Replace: Clear current list first",
            font=("Arial", 9),
            fg="black",
            anchor="w"
        )
        option2_label.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        # Button frame
        button_frame = tk.Frame(merge_window)
        button_frame.pack(pady=10)
        
        # Append button
        append_button = tk.Button(
            button_frame,
            text="Append",
            command=lambda: (
                result.__setitem__(0, True),
                merge_window.destroy()
            ),
            width=12,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10)
        )
        append_button.grid(row=0, column=0, padx=3)
        
        # Replace button
        replace_button = tk.Button(
            button_frame,
            text="Replace",
            command=lambda: (
                result.__setitem__(0, False),
                merge_window.destroy()
            ),
            width=12,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10)
        )
        replace_button.grid(row=0, column=1, padx=3)
        
        # Cancel button
        cancel_button = tk.Button(
            button_frame,
            text="Cancel",
            command=merge_window.destroy,
            width=12,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10)
        )
        cancel_button.grid(row=0, column=2, padx=3)
        
        # Wait for window to close and return result
        self.root.wait_window(merge_window)
        return result[0]
    
    def _parse_page_range(self, range_text: str, total_pages: int) -> List[int]:
        """Parse page range string into a list of zero-based page indices."""
        text = (range_text or "").strip().lower()
        if text == "" or text == "all":
            return list(range(total_pages))
        
        indices = set()
        parts = [part.strip() for part in text.split(",") if part.strip()]
        for part in parts:
            if "-" in part:
                start_str, end_str = [p.strip() for p in part.split("-", 1)]
                if not start_str or not end_str:
                    raise ValueError("Invalid range format")
                try:
                    start = int(start_str)
                    end = int(end_str)
                except ValueError:
                    raise ValueError("Page numbers must be integers")
                if start < 1 or end < 1 or start > end:
                    raise ValueError("Invalid range order")
                if end > total_pages:
                    raise ValueError(f"Page range exceeds total pages, file has {total_pages} pages")
                for page_num in range(start, end + 1):
                    indices.add(page_num - 1)
            else:
                try:
                    page_num = int(part)
                except ValueError:
                    raise ValueError("Page numbers must be integers")
                if page_num < 1 or page_num > total_pages:
                    raise ValueError(f"Page number out of range, file has {total_pages} pages")
                indices.add(page_num - 1)
        
        if not indices:
            raise ValueError("No valid pages selected")
        
        return sorted(indices)


    def _validate_page_range(self, index: int, var: tk.StringVar, entry_widget: tk.Entry):
        """Validate a page range after entry and revert on error."""
        if not (0 <= index < len(self.pdf_files)):
            return

        file_path = self.get_file_path(self.pdf_files[index])
        text = var.get().strip()
        normalized = text if text else "All"

        try:
            with open(file_path, 'rb') as pdf_file:
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                total_pages = len(pdf_reader.pages)
                self._parse_page_range(normalized, total_pages)
        except Exception as e:
            last_valid = self.page_range_last_valid.get(index, "All")
            var.set(last_valid)
            entry_widget.focus_set()
            entry_widget.selection_range(0, tk.END)
            
            # Get total pages for error message
            try:
                with open(file_path, 'rb') as pdf_file:
                    pdf_reader = PyPDF2.PdfReader(pdf_file)
                    total_pages = len(pdf_reader.pages)
                    pages_info = f"(Total pages: {total_pages})"
            except:
                pages_info = ""
            
            error_msg = (
                f"{Path(file_path).name} {pages_info}\n\n"
                f"Error: {e}\n\n"
                f"Valid formats:\n"
                f"  • All pages: 'All' or leave blank\n"
                f"  • Single page: '5'\n"
                f"  • Range: '1-10'\n"
                f"  • Multiple ranges: '1-3,5,7-9'"
            )
            self.show_error_dialog("Invalid Page Range", error_msg)
            return

        self.page_range_last_valid[index] = normalized
        self.set_page_range(index, normalized)

    def open_output_file(self):
        """Open the last combined PDF using the system default application"""
        path = self.last_output_file
        if not path:
            messagebox.showwarning("No file", "No combined PDF available to open.")
            return

        if not os.path.exists(path):
            self.show_error_dialog("Error", f"File not found: {path}")
            return

        try:
            if os.name == 'nt':
                os.startfile(path)
            else:
                webbrowser.open_new(path)
        except Exception as e:
            self.show_error_dialog("Error", f"Could not open file: {e}")

    def _is_page_blank(self, page) -> bool:
        """Detect if a PDF page is blank by checking both text and visual content."""
        try:
            # Check for text content
            text = page.extract_text()
            if text and text.strip():
                return False  # Has text, not blank
            
            # Check for visual content (images, graphics, etc.)
            # Check if page has resources with images or other content
            if "/Resources" in page:
                resources = page["/Resources"]
                
                # Check for images (XObjects)
                if "/XObject" in resources:
                    xobjects = resources["/XObject"]
                    if xobjects and len(xobjects) > 0:
                        return False  # Has images
                
                # Check for other content streams
                if "/Font" in resources and len(resources["/Font"]) > 0:
                    # Has fonts, might have invisible text or form fields
                    return False
            
            # If no text and no visual resources, consider it blank
            return True
            
        except Exception as e:
            # If we can't analyze the page, assume not blank to be safe
            return False
    
    def _image_to_pdf(self, image_path: str) -> str:
        """Convert an image file to a temporary PDF and return the path to the PDF.
        
        Supports: JPG, PNG, BMP, GIF, TIFF, etc.
        Returns the path to the temporary PDF file.
        """
        try:
            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import letter
            from reportlab.lib.utils import ImageReader
            import tempfile
            
            # Load the image
            img = Image.open(image_path)
            
            # Convert RGBA to RGB if needed (JPEGs don't support transparency)
            if img.mode in ('RGBA', 'LA', 'P'):
                # Create white background
                background = Image.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'P':
                    img = img.convert('RGBA')
                background.paste(img, mask=img.split()[-1] if len(img.split()) > 3 else None)
                img = background
            elif img.mode != 'RGB':
                img = img.convert('RGB')
            
            # Get image dimensions
            img_width, img_height = img.size
            
            # Create a temporary PDF file
            temp_fd, temp_pdf_path = tempfile.mkstemp(suffix='.pdf')
            os.close(temp_fd)  # Close the file descriptor, we'll use reportlab
            
            # Create PDF with same aspect ratio as image
            # Use a reasonable DPI (96) to scale the image
            dpi = 96
            page_width = (img_width / dpi) * 72  # Convert pixels to points (72 DPI in PDF)
            page_height = (img_height / dpi) * 72
            
            # Create canvas and add image
            c = canvas.Canvas(temp_pdf_path, pagesize=(page_width, page_height))
            c.drawImage(ImageReader(img), 0, 0, width=page_width, height=page_height)
            c.save()
            
            return temp_pdf_path
            
        except Exception as e:
            raise Exception(f"Failed to convert image to PDF: {str(e)}")
    
    def _scale_page(self, page, target_width: float, target_height: float):
        """Scale a page to fit target dimensions while maintaining aspect ratio."""
        try:
            box = page.mediabox
            current_width = float(box.width)
            current_height = float(box.height)
            
            # Calculate scale factors to fit within target dimensions
            scale_x = target_width / current_width
            scale_y = target_height / current_height
            scale = min(scale_x, scale_y)  # Use smaller scale to fit within target
            
            # Calculate new dimensions after scaling
            new_width = current_width * scale
            new_height = current_height * scale
            
            # Calculate centering offsets
            x_offset = (target_width - new_width) / 2
            y_offset = (target_height - new_height) / 2
            
            # Apply scaling and translation to center the content
            if scale != 1.0 or x_offset != 0 or y_offset != 0:
                # Scale and translate the content
                page.scale_by(scale)
                page.add_transform_matrix([1, 0, 0, 1, x_offset, y_offset])
                
            # Set mediabox to the full target size (not cropped)
            page.mediabox.lower_left = (0, 0)
            page.mediabox.upper_right = (target_width, target_height)
            
        except Exception as e:
            # If scaling fails, leave page as is
            pass
    
    def _add_watermark(self, page, text: str, opacity: float, font_size: int = 50, rotation: int = 45, position: str = "center"):
        """Add text watermark to a PDF page with optional Safe Mode to prevent clipping."""
        try:
            from reportlab.pdfgen import canvas
            from io import BytesIO
            import math
            
            # Get page dimensions
            box = page.mediabox
            width = float(box.width)
            height = float(box.height)
            
            # Normalize position string to lowercase for comparison
            position = position.lower()
            
            # Calculate if watermark would be clipped (Safe Mode detection)
            adjusted_font_size = font_size
            adjusted_position = position
            
            if self.watermark_safe_mode.get():
                # Estimate text width (approximate: 0.6 pixels per point per character)
                text_width_approx = len(text) * font_size * 0.5
                text_height_approx = font_size * 1.2
                
                # Calculate corners of rotated text bounding box
                angle_rad = math.radians(rotation)
                cos_a = abs(math.cos(angle_rad))
                sin_a = abs(math.sin(angle_rad))
                
                # Rotated bounding box dimensions
                rotated_width = text_width_approx * cos_a + text_height_approx * sin_a
                rotated_height = text_width_approx * sin_a + text_height_approx * cos_a
                
                # Safe margins (padding from edges)
                safe_margin = 40
                
                # Check if would clip at top/bottom positions
                if position in ["top", "bottom"]:
                    # At top/bottom, text is more likely to clip due to limited vertical space
                    # Check if rotated height exceeds available space
                    if position == "top" and rotated_height * 0.5 > height * 0.15 + safe_margin:
                        # Would clip at top - either reduce font or move to center
                        # Try to reduce font size first
                        max_font = int(font_size * (height * 0.15 / (rotated_height * 0.5)) * 0.95)
                        if max_font >= 10:
                            adjusted_font_size = max(10, max_font)
                        else:
                            # Font too large, move to center (safest for rotation)
                            adjusted_position = "center"
                            
                    elif position == "bottom" and rotated_height * 0.5 > height * 0.15 + safe_margin:
                        # Would clip at bottom - either reduce font or move to center
                        max_font = int(font_size * (height * 0.15 / (rotated_height * 0.5)) * 0.95)
                        if max_font >= 10:
                            adjusted_font_size = max(10, max_font)
                        else:
                            # Font too large, move to center
                            adjusted_position = "center"
                
                # Check if rotated width exceeds page width
                if rotated_width > width - safe_margin * 2 and adjusted_font_size > 10:
                    # Reduce font size to fit horizontally
                    max_font = int(adjusted_font_size * (width - safe_margin * 2) / rotated_width * 0.95)
                    adjusted_font_size = max(10, max_font)
            
            # Create watermark in memory
            packet = BytesIO()
            c = canvas.Canvas(packet, pagesize=(width, height))
            c.setFillAlpha(opacity)
            c.setFont("Helvetica-Bold", adjusted_font_size)
            c.setFillGray(0.5)
            
            # Draw watermark at specified position
            c.saveState()
            
            if adjusted_position == "top":
                # Position at top of page
                c.translate(width / 2, height * 0.85)
            elif adjusted_position == "bottom":
                # Position at bottom of page
                c.translate(width / 2, height * 0.15)
            else:  # center (default)
                # Position at center of page
                c.translate(width / 2, height / 2)
            
            c.rotate(rotation)
            c.drawCentredString(0, 0, text)
            c.restoreState()
            c.save()
            
            # Move to the beginning of the BytesIO buffer
            packet.seek(0)
            
            # Read the watermark PDF from memory
            watermark_pdf = PyPDF2.PdfReader(packet)
            watermark_page = watermark_pdf.pages[0]
            
            # Merge watermark with page
            page.merge_page(watermark_page)
            
        except ImportError:
            # reportlab not available, skip watermarking
            pass
        except Exception:
            # If watermarking fails, continue without it
            pass
    
    def _create_page_with_filename(self, filename: str, width: float, height: float):
        """Create a blank page with filename text centered."""
        try:
            from reportlab.pdfgen import canvas
            from io import BytesIO
            
            # Ensure dimensions are float
            width = float(width)
            height = float(height)
            
            # Scale text and spacing based on page height (792 = letter size height)
            scale_factor = height / 792.0
            base_font_size = 14
            scaled_font_size = int(base_font_size * scale_factor)
            scaled_line_height = int(18 * scale_factor)
            scaled_spacing_below_file = int(35 * scale_factor)
            scaled_line_spacing = int(20 * scale_factor)
            scaled_margin = int(50 * scale_factor)
            scaled_line_width = max(1, int(2 * scale_factor))
            
            # Create blank page with filename in the center
            packet = BytesIO()
            c = canvas.Canvas(packet, pagesize=(width, height))
            
            # Wrap filename if longer than 25 characters
            max_chars = 25
            filename_lines = []
            if len(filename) > max_chars:
                # Split filename into chunks of max_chars
                for i in range(0, len(filename), max_chars):
                    filename_lines.append(filename[i:i + max_chars])
            else:
                filename_lines = [filename]
            
            # Calculate total height needed for text block
            filename_height = len(filename_lines) * scaled_line_height
            total_text_height = 40 * scale_factor + filename_height + 15 * scale_factor + 15 * scale_factor
            
            # Draw filename and "follows" text centered vertically
            vertical_center = height / 2
            current_y = vertical_center + (total_text_height / 2) - 10 * scale_factor
            
            c.setFont("Helvetica", scaled_font_size)
            c.setFillGray(0.3)
            c.drawCentredString(width / 2, current_y, "File")
            
            # Draw filename lines with horizontal lines above and below
            c.setFont("Helvetica-Bold", scaled_font_size)
            current_y -= scaled_spacing_below_file
            
            # Draw horizontal line above filename
            c.setStrokeGray(0.3)
            c.setLineWidth(scaled_line_width)
            c.line(scaled_margin, current_y + scaled_line_spacing, width - scaled_margin, current_y + scaled_line_spacing)
            
            # Draw filename lines
            for line in filename_lines:
                c.drawCentredString(width / 2, current_y, line)
                current_y -= scaled_line_height
            
            # Draw horizontal line below filename
            c.line(scaled_margin, current_y + scaled_line_height - 10 * scale_factor, width - scaled_margin, current_y + scaled_line_height - 10 * scale_factor)
            
            # Draw "follows" text
            c.setFont("Helvetica", scaled_font_size)
            current_y -= 10 * scale_factor
            c.drawCentredString(width / 2, current_y, "follows")
            
            c.save()
            
            # Move to the beginning of the BytesIO buffer
            packet.seek(0)
            
            # Read the page from memory
            page_pdf = PyPDF2.PdfReader(packet)
            if len(page_pdf.pages) > 0:
                return page_pdf.pages[0]
            else:
                # If no pages were created, fall back to blank page
                blank_page = PyPDF2.PdfWriter().add_blank_page(width=width, height=height)
                return blank_page
            
        except ImportError:
            # reportlab not available, create blank page without text
            blank_page = PyPDF2.PdfWriter().add_blank_page(width=float(width), height=float(height))
            return blank_page
        except Exception as e:
            # If page creation fails, return blank page
            # This is a fallback - the page will have text but if it fails, at least we get a blank page
            blank_page = PyPDF2.PdfWriter().add_blank_page(width=float(width), height=float(height))
            return blank_page
    
    def _insert_toc_page(self, pdf_path: str, toc_entries: List[Dict]):
        """Insert Table of Contents pages at the beginning of the PDF using PyMuPDF.
        Creates multiple pages if needed to fit all entries."""
        try:
            # Open the PDF with PyMuPDF
            doc = fitz.open(pdf_path)
            
            # Use standard letter size for TOC pages
            page_width = 612  # Letter width
            page_height = 792  # Letter height
            
            # Set up fonts and positions
            title_font_size = 24
            entry_font_size = 11
            margin_left = 72  # 1 inch
            margin_top = 72   # 1 inch
            margin_bottom = 72  # 1 inch
            line_height = 20
            
            # Calculate how many entries fit per page
            available_height = page_height - margin_top - margin_bottom - (title_font_size + 25)
            entries_per_page = int(available_height / line_height)
            
            # Ensure at least 1 entry per page to avoid infinite loops
            if entries_per_page < 1:
                entries_per_page = 1
            
            # Split entries into pages
            toc_pages_data = []
            for i in range(0, len(toc_entries), entries_per_page):
                chunk = toc_entries[i:i + entries_per_page]
                page_number = len(toc_pages_data) + 1
                toc_pages_data.append({
                    'entries': chunk,
                    'page_number': page_number,
                    'total_pages': (len(toc_entries) + entries_per_page - 1) // entries_per_page
                })
            
            # Insert TOC pages in reverse order (so they end up at the beginning)
            for toc_page_idx in range(len(toc_pages_data) - 1, -1, -1):
                toc_page_info = toc_pages_data[toc_page_idx]
                
                # Insert a new page at the beginning
                toc_page = doc.new_page(0, width=page_width, height=page_height)
                
                # Add title
                title_text = "Table of Contents"
                if toc_page_info['total_pages'] > 1:
                    title_text += f" (Page {toc_page_info['page_number']} of {toc_page_info['total_pages']})"
                
                toc_page.insert_text(
                    (margin_left, margin_top + 20),
                    title_text,
                    fontsize=title_font_size,
                    fontname="helv",
                    color=(0, 0, 0)
                )
                
                # Draw a line under the title
                line_y = margin_top + title_font_size + 15
                toc_page.draw_line(
                    (margin_left, line_y),
                    (page_width - margin_left, line_y),
                    color=(0, 0, 0),
                    width=1
                )
                
                # Add TOC entries for this page
                current_y = line_y + 25
                max_filename_length = 80
                
                for entry in toc_page_info['entries']:
                    filename = entry['filename']
                    # +1 for each TOC page we're inserting, +1 for destination page offset
                    page_num = entry['page'] + len(toc_pages_data) + 1
                    
                    # Truncate long filenames
                    if len(filename) > max_filename_length:
                        filename = filename[:max_filename_length-3] + "..."
                    
                    # Create text for entry
                    entry_text = f"{filename}"
                    page_text = f"Page {page_num}"
                    
                    # Position for the entry text
                    text_rect = fitz.Rect(margin_left, current_y, page_width - 150, current_y + line_height)
                    
                    # Insert the filename text
                    toc_page.insert_textbox(
                        text_rect,
                        entry_text,
                        fontsize=entry_font_size,
                        fontname="helv",
                        color=(0, 0, 1),  # Blue color for clickable look
                        align=fitz.TEXT_ALIGN_LEFT
                    )
                    
                    # Insert the page number (right-aligned)
                    page_rect = fitz.Rect(page_width - 150, current_y, page_width - margin_left, current_y + line_height)
                    toc_page.insert_textbox(
                        page_rect,
                        page_text,
                        fontsize=entry_font_size,
                        fontname="helv",
                        color=(0.3, 0.3, 0.3),
                        align=fitz.TEXT_ALIGN_RIGHT
                    )
                    
                    # Create clickable link to the destination page
                    # Account for all TOC pages that will be inserted
                    dest_page_index = entry['page'] + len(toc_pages_data)
                    
                    # Validate page index is within bounds
                    if dest_page_index < len(doc):
                        link_rect = fitz.Rect(margin_left, current_y, page_width - margin_left, current_y + line_height)
                        link = {
                            "kind": fitz.LINK_GOTO,
                            "from": link_rect,
                            "page": dest_page_index,
                            "to": fitz.Point(0, 0),
                            "zoom": 0
                        }
                        toc_page.insert_link(link)
                    
                    current_y += line_height
            
            # Shift existing bookmarks to account for inserted TOC pages
            if toc_pages_data:
                try:
                    existing_toc = doc.get_toc()
                    if existing_toc:
                        page_offset = len(toc_pages_data)
                        for entry in existing_toc:
                            if len(entry) >= 3 and isinstance(entry[2], int):
                                entry[2] += page_offset
                        doc.set_toc(existing_toc)
                except Exception:
                    pass

            # Save to a temporary file, then replace the original
            import tempfile
            temp_fd, temp_path = tempfile.mkstemp(suffix='.pdf')
            os.close(temp_fd)
            
            try:
                doc.save(temp_path, garbage=4, deflate=True)
                doc.close()
                
                # Replace the original file with the temp file
                os.replace(temp_path, pdf_path)
                
            except Exception as save_error:
                doc.close()
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                raise save_error
            
        except Exception as e:
            print(f"Warning: Could not insert TOC page: {e}")
    
    def _compress_page(self, page, compression_level: str):
        """Apply compression to a PDF page based on compression level."""
        try:
            # Compression mapping - JPEG quality values
            quality_map = {
                "Low": 95,      # Minimal compression
                "Medium": 75,   # Moderate compression
                "High": 50,     # High compression
                "Maximum": 30   # Maximum compression
            }
            
            quality = quality_map.get(compression_level, 75)
            
            # Compress images in the page
            if '/Resources' in page and '/XObject' in page['/Resources']:
                xobjects = page['/Resources']['/XObject'].get_object()
                for obj_name in xobjects:
                    obj = xobjects[obj_name]
                    if obj.get('/Subtype') == '/Image':
                        try:
                            # Get image data
                            if hasattr(obj, 'get_data'):
                                image_data = obj.get_data()
                                width = obj.get('/Width', 0)
                                height = obj.get('/Height', 0)
                                
                                # Only compress if we have valid dimensions
                                if width > 0 and height > 0:
                                    # Try to load and recompress the image
                                    try:
                                        from PIL import Image
                                        import io
                                        
                                        # Attempt to load image from data
                                        img = Image.open(io.BytesIO(image_data))
                                        
                                        # Convert to RGB if necessary (for JPEG compression)
                                        if img.mode in ('RGBA', 'LA', 'P'):
                                            # Create white background for transparency
                                            background = Image.new('RGB', img.size, (255, 255, 255))
                                            if img.mode == 'P':
                                                img = img.convert('RGBA')
                                            background.paste(img, mask=img.split()[-1] if img.mode in ('RGBA', 'LA') else None)
                                            img = background
                                        elif img.mode != 'RGB':
                                            img = img.convert('RGB')
                                        
                                        # Compress image to JPEG at specified quality
                                        output = io.BytesIO()
                                        img.save(output, format='JPEG', quality=quality, optimize=True)
                                        compressed_data = output.getvalue()
                                        
                                        # Only replace if compression actually reduced size
                                        if len(compressed_data) < len(image_data):
                                            # Update the image object with compressed data
                                            obj._data = compressed_data
                                            obj[PyPDF2.generic.NameObject('/Filter')] = PyPDF2.generic.NameObject('/DCTDecode')
                                            obj[PyPDF2.generic.NameObject('/ColorSpace')] = PyPDF2.generic.NameObject('/DeviceRGB')
                                            if '/DecodeParms' in obj:
                                                del obj['/DecodeParms']
                                    except Exception:
                                        # If PIL compression fails, try flate encoding
                                        if hasattr(obj, 'flate_encode'):
                                            obj.flate_encode()
                            else:
                                # Fallback to flate encoding
                                if hasattr(obj, 'flate_encode'):
                                    obj.flate_encode()
                        except Exception:
                            pass
        except Exception:
            # If compression fails, continue without it
            pass


if __name__ == "__main__":
    _enable_dpi_awareness()
    root = tk.Tk()
    # Show splash screen before main app loads
    splash_path = os.path.join(os.path.dirname(__file__), "images", "splashscreen.png")
    show_splash(root, splash_path)
    app = PDFCombinerApp(root)
    root.mainloop()
