import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import sys
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

__VERSION__ = "1.2.1"

class PDFCombinerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Combiner")
        
        # Center window horizontally and align to top
        window_width = 700
        window_height = 600
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
        self.output_directory = str(Path.home() / "Documents")
        self.add_files_directory = str(Path.home() / "Documents")  # Default directory for adding files
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
        self.insert_blank_pages = tk.BooleanVar(value=False)  # Insert blank pages between files
        self.rotation_vars = {}  # Map of index to tk.StringVar for rotation dropdowns
        self.page_range_vars = {}  # Map of index to tk.StringVar for page ranges
        self.page_range_last_valid = {}  # Track last valid page range per index
        
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
        style.configure('TNotebook.Tab', padding=[10, 4], font=('Arial', 10, 'bold'), background='#D0D0D0', foreground='#333333', focuscolor='#D0D0D0')
        style.map('TNotebook.Tab', 
                  background=[('selected', '#4A90E2'), ('active', '#5B9FE8')],
                  foreground=[('selected', 'white'), ('active', 'white')],
                  padding=[('selected', [10, 4])])
        
        # Configure Combobox style to match file list background
        style.configure('TCombobox', fieldbackground='white', background='white', foreground='black', relief='flat', borderwidth=0)
        style.map('TCombobox',
                  fieldbackground=[('readonly', 'white'), ('disabled', 'white')],
                  background=[('readonly', 'white'), ('disabled', 'white')])
        
        # Create notebook (tabbed interface)
        self.notebook = ttk.Notebook(root, style='TNotebook')
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # ===== INPUT TAB =====
        input_frame = tk.Frame(self.notebook)
        self.notebook.add(input_frame, text="Input")
        
        # Spacer to move content down by 25 pixels
        spacer_frame = tk.Frame(input_frame, height=25)
        spacer_frame.pack()
        spacer_frame.pack_propagate(False)
        
        # Title and preview checkbox frame - same line
        title_frame = tk.Frame(input_frame)
        title_frame.pack(anchor=tk.W, fill=tk.X, padx=10, pady=(5, 5))
        
        # Title above list
        title_label = tk.Label(title_frame, text="List and Order of Files to Combine:", font=("Arial", 10, "bold"))
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
        self.filename_hdr = tk.Label(header_frame, text="Filename", font=hdr_font, bg="#E0E0E0", width=58, anchor='w')
        self.filename_hdr.pack(side=tk.LEFT)
        self.filename_hdr.bind("<Button-1>", lambda e: self.on_sort_clicked('name'))
        self.filename_hdr.bind("<Enter>", lambda e: self.filename_hdr.config(cursor="hand2"))
        self.filename_hdr.bind("<Leave>", lambda e: self.filename_hdr.config(cursor="arrow"))

        # File Size header - clickable
        self.size_hdr = tk.Label(header_frame, text="Size", font=hdr_font, bg="#E0E0E0", width=10, anchor='w')
        self.size_hdr.pack(side=tk.LEFT)
        self.size_hdr.bind("<Button-1>", lambda e: self.on_sort_clicked('size'))
        self.size_hdr.bind("<Enter>", lambda e: self.size_hdr.config(cursor="hand2"))
        self.size_hdr.bind("<Leave>", lambda e: self.size_hdr.config(cursor="arrow"))

        # Date header - clickable
        self.date_hdr = tk.Label(header_frame, text="Date", font=hdr_font, bg="#E0E0E0", width=11, anchor='w')
        self.date_hdr.pack(side=tk.LEFT)
        self.date_hdr.bind("<Button-1>", lambda e: self.on_sort_clicked('date'))
        self.date_hdr.bind("<Enter>", lambda e: self.date_hdr.config(cursor="hand2"))
        self.date_hdr.bind("<Leave>", lambda e: self.date_hdr.config(cursor="arrow"))
        
        pages_hdr = tk.Label(header_frame, text="Pages", font=hdr_font, bg="#E0E0E0", width=10, anchor='w')
        pages_hdr.pack(side=tk.LEFT, padx=(4, 0))
        
        rot_hdr = tk.Label(header_frame, text="Rotate", font=hdr_font, bg="#E0E0E0", width=6, anchor='c')
        rot_hdr.pack(side=tk.LEFT, padx=2)
        
        # Sub-frame for custom list frame and scrollbar (sized for ~11 rows)
        listbox_scroll_frame = tk.Frame(list_frame, height=270)
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
            
            # Update the scrollregion to encompass the frame
            self.file_list_frame.update_idletasks()
            # Use the frame's required size for scrollregion
            frame_width = self.file_list_frame.winfo_reqwidth()
            frame_height = self.file_list_frame.winfo_reqheight()
            
            # Get canvas height to prevent scrolling above content
            canvas_height = self.file_list_canvas.winfo_height()
            if canvas_height <= 1:
                canvas_height = 270  # Default height fallback
            
            # Ensure scrollregion is at least as tall as canvas to prevent blank lines when scrolling
            scrollregion_height = max(frame_height, canvas_height)
            new_region = (0, 0, max(frame_width, 1), scrollregion_height)
            
            if new_region != self.last_scrollregion:
                self.last_scrollregion = new_region
                self.file_list_canvas.configure(scrollregion=new_region)
                # Make canvas window width match canvas width
                if self.file_list_canvas.winfo_width() > 1:
                    self.file_list_canvas.itemconfig(canvas_window, width=self.file_list_canvas.winfo_width())
        
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
        
        # File count label
        self.count_label = tk.Label(input_frame, text="Files to combine: 0", font=("Arial", 9))
        self.count_label.pack(pady=1)
        
        # Drag and drop instruction
        drag_drop_note = tk.Label(
            input_frame,
            text="After adding files, single click to select a file. Ctrl-Click to select multiple files. Click and drag files to reorder.\nHover to preview the first page. Double-click to open a file. Click column headers to sort.",
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
            text="Add PDFs to Combine...",
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
            text="Clear All",
            command=self.clear_files,
            width=18,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10),
            state=tk.DISABLED  # Start disabled since list is empty
        )
        self.clear_button.grid(row=0, column=2, padx=5)
        
        # ===== OUTPUT TAB =====
        output_frame_main = tk.Frame(self.notebook)
        self.notebook.add(output_frame_main, text="Output")
        
        # Padding frame for better spacing
        output_content_frame = tk.Frame(output_frame_main)
        output_content_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # Output settings frame
        output_frame = tk.LabelFrame(output_content_frame, text="Output Settings", font=("Arial", 10, "bold"), padx=10, pady=8)
        output_frame.pack(pady=5, fill=tk.X)
        
        # Filename frame
        filename_frame = tk.Frame(output_frame)
        filename_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(filename_frame, text="Filename for combined PDF:", font=("Arial", 9)).pack(side=tk.LEFT, padx=5)
        
        filename_entry = tk.Entry(filename_frame, textvariable=self.output_filename, font=("Arial", 9), width=30)
        filename_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        #tk.Label(filename_frame, text=".pdf", font=("Arial", 9)).pack(side=tk.LEFT)
        
        # Location frame (boxed to highlight save location)
        location_frame = tk.Frame(output_frame)
        location_frame.pack(fill=tk.X, pady=5)
        
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
            padx=6,
            pady=4,
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
        options_frame.pack(pady=(5, 0), padx=0, fill=tk.X)

        bookmark_checkbox = tk.Checkbutton(
            options_frame,
            text="Add filename bookmarks to the combined PDF",
            variable=self.add_filename_bookmarks,
            command=self._save_settings,
            font=("Arial", 9)
        )
        bookmark_checkbox.pack(anchor="w", padx=8, pady=(4, 2))

        blank_pages_checkbox = tk.Checkbutton(
            options_frame,
            text="Insert blank page between files",
            variable=self.insert_blank_pages,
            command=self._save_settings,
            font=("Arial", 9)
        )
        blank_pages_checkbox.pack(anchor="w", padx=8, pady=(0, 4))
        
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
        bottom_frame = tk.Frame(root, height=40)
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
        
        # Help button
        help_button = tk.Button(
            center_frame,
            text="Help",
            command=self.show_help,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 11),
            height=1,
            width=13
        )
        help_button.pack(side=tk.LEFT, padx=5)
        
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
        
        # Copyright notice aligned to the right in the same row
        copyright_label = tk.Label(
            bottom_frame,
            text="© 2026 tgtechy",
            font=("Arial", 8, "underline"),
            fg="#1A5FB4",
            bg=bottom_frame.cget("bg"),
            cursor="hand2"
        )
        copyright_label.place(relx=1.0, rely=0.5, anchor="e", x=-5)
        copyright_label.bind(
            "<Button-1>",
            lambda e: webbrowser.open_new("https://github.com/tgtechy/CombinePDFs")
        )
        
        # Version label aligned to the left
        version_label = tk.Label(
            bottom_frame,
            text=f"v{__VERSION__}",
            font=("Arial", 8),
            fg="#606060",
            bg=bottom_frame.cget("bg")
        )
        version_label.place(relx=0.0, rely=0.5, anchor="w", x=5)
        
        # Set up tab change handler to maintain focus on add button
        def on_tab_changed(event):
            if self.notebook.index(self.notebook.select()) == 0:  # Input tab
                self.add_button.focus_set()
        
        self.notebook.bind("<<NotebookTabChanged>>", on_tab_changed)
        
        # Set initial focus to add button
        self.add_button.focus_set()
    
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
                
                # Load preview enabled state
                if 'preview_enabled' in settings:
                    self.preview_enabled.set(settings['preview_enabled'])
                
                # Load add filename bookmarks state
                if 'add_filename_bookmarks' in settings:
                    self.add_filename_bookmarks.set(settings['add_filename_bookmarks'])
                
                # Load insert blank pages state
                if 'insert_blank_pages' in settings:
                    self.insert_blank_pages.set(settings['insert_blank_pages'])
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
                'preview_enabled': self.preview_enabled.get(),
                'add_filename_bookmarks': self.add_filename_bookmarks.get(),
                'insert_blank_pages': self.insert_blank_pages.get()
            }
            with open(self.config_file, 'w') as f:
                json.dump(settings, f, indent=2)
        except Exception:
            # Silently fail if we can't save settings
            pass
    
    def add_files(self):
        """Open file dialog to select PDF files"""
        files = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=self.add_files_directory
        )
        
        added_count = 0
        duplicate_count = 0
        duplicates = []
        non_pdf_count = 0
        non_pdf_files = []
        
        # Get existing paths for duplicate checking
        existing_paths = {entry['path'] for entry in self.pdf_files}
        
        for file in files:
            # Check if file is a PDF
            if not file.lower().endswith('.pdf'):
                non_pdf_count += 1
                non_pdf_files.append(Path(file).name)
                continue
            
            if file not in existing_paths:
                # Create dict entry with path and default rotation/page range
                self.pdf_files.append({'path': file, 'rotation': 0, 'page_range': 'All'})
                added_count += 1
                # Update the add files directory to the directory of the last selected file
                self.add_files_directory = str(Path(file).parent)
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
        
        # Save settings if files were added
        if added_count > 0:
            self._save_settings()
        
        # Show warning if non-PDF files were attempted
        if non_pdf_count > 0:
            non_pdf_text = "\n".join(f"  • {file}" for file in non_pdf_files)
            messagebox.showwarning(
                "Invalid Files",
                f"The following file(s) are not PDF files and were not added:\n\n{non_pdf_text}"
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

    def get_page_range(self, file_entry: Dict[str, any]) -> str:
        """Extract page range from entry dict"""
        return file_entry.get('page_range', 'All')
    
    def set_rotation(self, index: int, degrees: int):
        """Update rotation for a file at given index"""
        if 0 <= index < len(self.pdf_files):
            self.pdf_files[index]['rotation'] = degrees

    def set_page_range(self, index: int, page_range: str):
        """Update page range for a file at given index"""
        if 0 <= index < len(self.pdf_files):
            cleaned = page_range.strip()
            self.pdf_files[index]['page_range'] = cleaned if cleaned else 'All'
    
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
            max_filename_len = 55
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
        self.row_visual_state.clear()  # Clear cached visual state since rows are rebuilt
        
        # Update button states after clearing
        self.root.after_idle(self._update_button_states)
        
        for i, pdf_entry in enumerate(self.pdf_files):
            file_path = self.get_file_path(pdf_entry)
            rotation = self.get_rotation(pdf_entry)
            
            # Create row frame - disable focus to prevent focus-change flicker
            row_frame = tk.Frame(self.file_list_frame, bg="white", takefocus=0)
            row_frame.pack(fill=tk.X, padx=0, pady=0, anchor='nw')
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
            num_label = tk.Label(row_frame, text=f"{i+1}", font=("Consolas", 8), bg="white", width=4, anchor='e')
            num_label.pack(side=tk.LEFT, padx=(0, 2), pady=0, ipady=0, anchor='nw')
            num_label.bind("<Button-1>", lambda e, idx=i: self.on_row_click(e, idx))
            num_label.bind("<B1-Motion>", lambda e, idx=i: self.on_row_drag(e, idx))
            num_label.bind("<ButtonRelease-1>", lambda e, idx=i: self.on_row_release(e, idx))
            num_label.bind("<Motion>", lambda e, idx=i: self.on_row_hover(e, idx))
            num_label.bind("<Leave>", self.on_row_leave)
            num_label.bind("<Double-Button-1>", lambda e, idx=i: self.on_row_double_click(e, idx))
            
            # Filename label
            filename_label = tk.Label(row_frame, text=filename, font=("Consolas", 8), bg="white", width=58, anchor='w', justify=tk.LEFT)
            filename_label.pack(side=tk.LEFT, pady=0, ipady=0, anchor='nw')
            filename_label.bind("<Button-1>", lambda e, idx=i: self.on_row_click(e, idx))
            filename_label.bind("<B1-Motion>", lambda e, idx=i: self.on_row_drag(e, idx))
            filename_label.bind("<ButtonRelease-1>", lambda e, idx=i: self.on_row_release(e, idx))
            filename_label.bind("<Motion>", lambda e, idx=i: self.on_row_hover(e, idx))
            filename_label.bind("<Leave>", self.on_row_leave)
            filename_label.bind("<Double-Button-1>", lambda e, idx=i: self.on_row_double_click(e, idx))
            
            # File size label
            size_label = tk.Label(row_frame, text=size_str, font=("Consolas", 8), bg="white", width=10, anchor='w')
            size_label.pack(side=tk.LEFT, pady=0, ipady=0, anchor='nw')
            size_label.bind("<Button-1>", lambda e, idx=i: self.on_row_click(e, idx))
            size_label.bind("<B1-Motion>", lambda e, idx=i: self.on_row_drag(e, idx))
            size_label.bind("<ButtonRelease-1>", lambda e, idx=i: self.on_row_release(e, idx))
            size_label.bind("<Motion>", lambda e, idx=i: self.on_row_hover(e, idx))
            size_label.bind("<Leave>", self.on_row_leave)
            size_label.bind("<Double-Button-1>", lambda e, idx=i: self.on_row_double_click(e, idx))
            
            # Date label
            date_label = tk.Label(row_frame, text=date_str, font=("Consolas", 8), bg="white", width=11, anchor='w')
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
                width=10,
                font=("Consolas", 8)
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
            rotation_var = tk.StringVar(value=str(rotation))
            self.rotation_vars[i] = rotation_var
            
            rotation_dropdown = ttk.Combobox(
                row_frame,
                textvariable=rotation_var,
                values=["0", "90", "180", "270"],
                width=4,
                state="readonly",
                font=("Consolas", 8)
            )
            rotation_dropdown.pack(side=tk.LEFT, padx=2, pady=0, ipady=0, anchor='nw')
            
            # Bind rotation change
            def on_rotation_change(var, idx=i):
                try:
                    degrees = int(var.get())
                    self.set_rotation(idx, degrees)
                except ValueError:
                    pass
            
            rotation_var.trace("w", lambda *args, var=rotation_var, idx=i: on_rotation_change(var, idx))
            
            # Bind events to dropdown too for consistency
            rotation_dropdown.bind("<Button-1>", lambda e, idx=i: self.on_row_click(e, idx))
            rotation_dropdown.bind("<Motion>", lambda e, idx=i: self.on_row_hover(e, idx))
            rotation_dropdown.bind("<Leave>", self.on_row_leave)
        
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
                messagebox.showerror("Error", f"Could not open file: {e}")
    
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
        """Show a preview popup with PDF thumbnail"""
        self.hide_preview()
        
        try:
            # Create preview window
            self.preview_window = tk.Toplevel(self.root)
            self.preview_window.wm_overrideredirect(True)
            self.preview_window.wm_attributes("-topmost", True)
            
            # Create main frame
            main_frame = tk.Frame(self.preview_window, bg="white", relief=tk.SOLID, borderwidth=1)
            main_frame.pack(padx=5, pady=5)
            
            # Get PDF info
            file_stat = os.stat(file_path)
            size_bytes = file_stat.st_size
            
            # Format file size
            if size_bytes < 1024:
                size_str = f"{size_bytes} B"
            elif size_bytes < 1024 * 1024:
                size_str = f"{size_bytes / 1024:.1f} KB"
            else:
                size_str = f"{size_bytes / (1024 * 1024):.1f} MB"
            
            # Get page count
            with open(file_path, 'rb') as pdf_file:
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                page_count = len(pdf_reader.pages)
            
            # Convert first page to image using PyMuPDF
            try:
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
    
    
    def show_help(self):
        """Display help dialog with program instructions"""
        help_text = """ADDING FILES
• Click "Add PDFs to Combine..." to select PDF files to combine
• Select one or multiple files from your computer using the file browser
• Only PDF files can be added; other file types will be rejected
• The same file cannot be added twice

ORGANIZING FILES
• Drag files up or down to change the order they'll be combined
• Click a file to select it, Ctrl+Click on other files to select
  more than one file
• Hover over a file row to see the full path in the status bar at
  the bottom of the screen

SORTING
• Click column headers (Filename, File Size, Date) to sort
• Click again to reverse the sort order (arrows show sort direction)
• An up arrow (▲) means ascending, down arrow (▼) means descending

FILE PROPERTIES
• Page rotation: Set 0°, 90°, 180°, or 270° (clockwise) for each file
• Pages: Specify which pages to include in the combined PDF using:
  - "All" or leave empty for all pages
  - Single page: "5" (without the quotes)
  - Range: "1-10"    (without the quotes)
  - Multiple ranges: "1-3,5,7-9" (without the quotes)

OUTPUT SETTINGS
• Enter the desired filename for the combined PDF
• Click "Browse" to choose where to save the combined PDF
• Check "Add filename bookmarks" to create PDF bookmarks
  from each source file's name in the combined PDF
• Existing bookmarks in files will be preserved

COMBINING PDFs
• At least 2 files are required to combine
• Click "Combine PDFs" to merge the files
• Review the summary and click "Proceed"
• The combined PDF will be created at your chosen location

PREVIEW
• Hover over a file to see a thumbnail of its first page
• Uncheck "Preview first page on hover" to disable previews

STATUS BAR
• The bottom status bar shows the full path of the file
  you're currently hovering over or have selected"""

        # Create help window
        help_window = tk.Toplevel(self.root)
        help_window.title("PDF Combiner - Help")
        help_window.geometry("550x600")
        help_window.transient(self.root)
        
        # Center the help window
        help_window.update_idletasks()
        window_width = 550
        window_height = 600
        screen_width = help_window.winfo_screenwidth()
        screen_height = help_window.winfo_screenheight()
        center_x = int((screen_width - window_width) / 2)
        center_y = int((screen_height - window_height) / 2)
        help_window.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        
        # Create a frame with scrollbar
        help_frame = tk.Frame(help_window)
        help_frame.pack(anchor=tk.W, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Scrollbar
        scrollbar = tk.Scrollbar(help_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Text widget
        text_widget = tk.Text(
            help_frame,
            font=("Arial", 9),
            wrap=tk.WORD,
            yscrollcommand=scrollbar.set,
            bg="white",
            fg="#000000",
            padx=10,
            pady=10
        )
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=text_widget.yview)
        
        # Insert help text
        text_widget.insert(tk.END, help_text)
        text_widget.config(state=tk.DISABLED)  # Make read-only
        
        # Close button
        close_button = tk.Button(
            help_window,
            text="Close",
            command=help_window.destroy,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10),
            width=15
        )
        close_button.pack(pady=10)
    
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
    
    def combine_pdfs(self):
        """Combine selected PDF files"""
        if len(self.pdf_files) < 2:
            messagebox.showerror("Error", "Please select at least 2 PDF files to combine.")
            return
        
        # Validate filename
        filename = self.output_filename.get().strip()
        if not filename:
            messagebox.showerror("Error", "Please enter an output filename.")
            return
        
        # Ensure filename ends with .pdf
        if not filename.lower().endswith(".pdf"):
            filename += ".pdf"
        
        # Full output path
        output_file = str(Path(self.output_directory) / filename)

        # If output exists, ask user whether to overwrite
        if os.path.exists(output_file):
            if not messagebox.askyesno("Overwrite File", f"'{output_file}' already exists. Overwrite?"):
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
                
                # Count pages
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
                        messagebox.showerror("Invalid Page Range", error_msg)
                        return
                
                # Get file size
                total_size_bytes += os.path.getsize(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read PDF information: {e}")
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
        summary_window.geometry("550x280")
        summary_window.resizable(False, False)
        summary_window.transient(self.root)
        summary_window.grab_set()
        
        # Center the summary window
        summary_window.update_idletasks()
        x = (summary_window.winfo_screenwidth() // 2) - (550 // 2)
        y = (summary_window.winfo_screenheight() // 2) - (280 // 2)
        summary_window.geometry(f"550x280+{x}+{y}")
        
        # Title
        title_label = tk.Label(
            summary_window,
            text="Combine Summary",
            font=("Arial", 12, "bold"),
            pady=10
        )
        title_label.pack()
        
        # Info frame
        info_frame = tk.Frame(summary_window)
        info_frame.pack(pady=10, padx=20, fill=tk.X)
        
        # Files count
        files_row = tk.Frame(info_frame)
        files_row.pack(fill=tk.X, pady=5)
        files_label = tk.Label(
            files_row,
            text="Files to combine:",
            font=("Arial", 10, "bold"),
            fg="black",
            anchor="w"
        )
        files_label.pack(side=tk.LEFT)
        files_value = tk.Label(
            files_row,
            text=f"  {len(files_to_combine)} files",
            font=("Arial", 10, "bold"),
            fg="#0066CC",
            anchor="w"
        )
        files_value.pack(side=tk.LEFT)
        
        # Total pages
        pages_row = tk.Frame(info_frame)
        pages_row.pack(fill=tk.X, pady=5)
        pages_label = tk.Label(
            pages_row,
            text="Total pages:",
            font=("Arial", 10, "bold"),
            fg="black",
            anchor="w"
        )
        pages_label.pack(side=tk.LEFT)
        if blank_pages > 0:
            pages_text = f"  {total_pages} pages ({original_pages} from PDFs + {blank_pages} blank separator pages)"
        else:
            pages_text = f"  {total_pages} pages"
        pages_value = tk.Label(
            pages_row,
            text=pages_text,
            font=("Arial", 10, "bold"),
            fg="#0066CC",
            anchor="w"
        )
        pages_value.pack(side=tk.LEFT)
        
        # Total size
        size_row = tk.Frame(info_frame)
        size_row.pack(fill=tk.X, pady=5)
        size_label = tk.Label(
            size_row,
            text="Total size:",
            font=("Arial", 10, "bold"),
            fg="black",
            anchor="w"
        )
        size_label.pack(side=tk.LEFT)
        size_value = tk.Label(
            size_row,
            text=f"  {size_str}",
            font=("Arial", 10, "bold"),
            fg="#0066CC",
            anchor="w"
        )
        size_value.pack(side=tk.LEFT)
        
        # Save path
        path_row = tk.Frame(info_frame)
        path_row.pack(fill=tk.X, pady=5)
        path_label = tk.Label(
            path_row,
            text="Save to:",
            font=("Arial", 10, "bold"),
            fg="black",
            anchor="w"
        )
        path_label.pack(side=tk.LEFT, anchor="nw")
        path_value = tk.Label(
            path_row,
            text=f"  {output_file}",
            font=("Arial", 9, "bold"),
            fg="#0066CC",
            anchor="w",
            wraplength=420,
            justify=tk.LEFT
        )
        path_value.pack(side=tk.LEFT, anchor="nw", fill=tk.X, expand=True)
        
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
        progress_window.geometry("400x190")
        progress_window.resizable(False, False)
        progress_window.transient(self.root)
        progress_window.grab_set()
        
        # Center the progress window
        progress_window.update_idletasks()
        x = (progress_window.winfo_screenwidth() // 2) - (400 // 2)
        y = (progress_window.winfo_screenheight() // 2) - (190 // 2)
        progress_window.geometry(f"400x190+{x}+{y}")
        
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
            try:
                # Create PDF writer object
                pdf_writer = PyPDF2.PdfWriter()
                
                # Track current page number for bookmarks
                current_page_num = 0
                
                # Add all PDFs to writer in the selected order with rotation applied
                for i, file_entry in enumerate(files_to_combine):
                    # Check if cancelled
                    if cancel_flag['cancelled']:
                        self.root.after(0, lambda: (
                            progress_window.destroy(),
                            messagebox.showinfo("Cancelled", "PDF combining operation was cancelled.")
                        ))
                        return
                    
                    file_path = self.get_file_path(file_entry)
                    rotation = self.get_rotation(file_entry)
                    page_range = self.get_page_range(file_entry)
                    
                    # Update progress
                    self.root.after(0, lambda idx=i, f=file_path: (
                        progress_label.config(text=f"Processing: {Path(f).name}"),
                        progress_bar.config(value=idx),
                        counter_label.config(text=f"{idx} of {len(files_to_combine)} files processed")
                    ))
                    
                    # Read the PDF and add pages with rotation
                    with open(file_path, 'rb') as pdf_file:
                        pdf_reader = PyPDF2.PdfReader(pdf_file)
                        total_file_pages = len(pdf_reader.pages)
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
                                messagebox.showerror("Invalid Page Range", error_msg)
                            ))
                            return
                        
                        # Add bookmark at the start of this file's pages (optional)
                        parent_bookmark = None
                        if self.add_filename_bookmarks.get() and len(page_indices) > 0:
                            bookmark_title = Path(file_path).stem  # Filename without extension
                            parent_bookmark = pdf_writer.add_outline_item(bookmark_title, current_page_num)
                        
                        # Always copy original bookmarks from source PDF (nested under file bookmark if it exists)
                        try:
                            # Use parent_bookmark if we added one, otherwise None for root level
                            self._copy_bookmarks(pdf_reader, pdf_writer, parent_bookmark, current_page_num, page_indices)
                        except Exception:
                            # If bookmark copying fails, continue without them
                            pass
                        
                        for page_index in page_indices:
                            page = pdf_reader.pages[page_index]
                            # Apply rotation if specified
                            if rotation != 0:
                                page.rotate(rotation)
                            pdf_writer.add_page(page)
                            current_page_num += 1
                        
                        # Insert blank page between files if enabled (but not after the last file)
                        if self.insert_blank_pages.get() and i < len(files_to_combine) - 1:
                            blank_page = pdf_writer.add_blank_page(width=pdf_reader.pages[0].mediabox.width if len(pdf_reader.pages) > 0 else 612,
                                                                     height=pdf_reader.pages[0].mediabox.height if len(pdf_reader.pages) > 0 else 792)
                            current_page_num += 1
                
                # Check if cancelled before writing
                if cancel_flag['cancelled']:
                    self.root.after(0, lambda: (
                        progress_window.destroy(),
                        messagebox.showinfo("Cancelled", "PDF combining operation was cancelled.")
                    ))
                    return
                
                # Update for writing phase
                self.root.after(0, lambda: (
                    progress_label.config(text="Writing combined PDF..."),
                    progress_bar.config(value=len(files_to_combine)),
                    counter_label.config(text=f"{len(files_to_combine)} of {len(files_to_combine)} files processed"),
                    cancel_button.config(state='disabled')
                ))
                
                # Write combined PDF
                with open(output_file, 'wb') as out_file:
                    pdf_writer.write(out_file)
                
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
                    messagebox.showerror("Error", error_msg)
                ))
            except Exception as e:
                error_msg = f"An error occurred: {str(e)}"
                self.root.after(0, lambda: (
                    progress_window.destroy(),
                    messagebox.showerror("Error", error_msg)
                ))
        
        # Start the thread
        thread = threading.Thread(target=combine_thread, daemon=True)
        thread.start()
    
    def show_success_dialog(self, output_file):
        """Show success dialog and ask to open the file"""
        if messagebox.askyesno("Success", f"PDFs combined successfully!\n\nOpen the combined PDF?\n\n{output_file}"):
            try:
                if os.name == 'nt':
                    os.startfile(output_file)
                else:
                    webbrowser.open_new(output_file)
            except Exception as e:
                messagebox.showerror("Error", f"Could not open file: {e}")

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

    def _copy_bookmarks(self, reader, writer, parent, page_offset, page_indices):
        """Recursively copy bookmarks from source PDF to combined PDF, adjusting page numbers."""
        try:
            outlines = reader.outline
            if not outlines:
                return
            
            self._copy_outline_items(reader, writer, outlines, parent, page_offset, page_indices)
        except Exception:
            # If reading outlines fails, silently continue without them
            pass
    
    def _copy_outline_items(self, reader, writer, items, parent, page_offset, page_indices):
        """Process outline items recursively."""
        for item in items:
            if isinstance(item, list):
                # Nested list of items
                self._copy_outline_items(reader, writer, item, parent, page_offset, page_indices)
            else:
                # Individual bookmark item
                try:
                    # Get the page number this bookmark points to
                    if hasattr(item, 'page'):
                        page_obj = item.page
                        if page_obj is not None:
                            # Get the index of this page in the source PDF
                            try:
                                source_page_num = reader.pages.index(page_obj)
                            except (ValueError, AttributeError):
                                continue
                            
                            # Check if this page is included in our selected range
                            if source_page_num not in page_indices:
                                continue
                            
                            # Calculate position in our selected pages
                            position_in_selection = page_indices.index(source_page_num)
                            new_page_num = page_offset + position_in_selection
                            
                            # Get bookmark title
                            title = item.get('/Title', 'Untitled')
                            
                            # Add bookmark as child of parent
                            new_bookmark = writer.add_outline_item(title, new_page_num, parent=parent)
                            
                            # Recursively process children if any
                            if hasattr(item, 'node') and hasattr(item.node, 'children'):
                                children = item.node.children
                                if children:
                                    self._copy_outline_items(reader, writer, children, new_bookmark, page_offset, page_indices)
                except Exception:
                    # Skip bookmarks that fail to process
                    continue

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
            messagebox.showerror("Invalid Page Range", error_msg)
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
            messagebox.showerror("Error", f"File not found: {path}")
            return

        try:
            if os.name == 'nt':
                os.startfile(path)
            else:
                webbrowser.open_new(path)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file: {e}")



if __name__ == "__main__":
    root = tk.Tk()
    app = PDFCombinerApp(root)
    root.mainloop()
