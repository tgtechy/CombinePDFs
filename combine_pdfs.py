import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import sys
import PyPDF2
from typing import List
from datetime import datetime
import os
import webbrowser
from PIL import Image, ImageTk
import io
import fitz  # PyMuPDF
import threading


class PDFCombinerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Combiner")
        
        # Center window on screen
        window_width = 700
        window_height = 580
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        center_x = int((screen_width - window_width) / 2)
        center_y = int((screen_height - window_height) / 2)
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
        
        self.pdf_files: List[str] = []
        self.combine_order = tk.StringVar(value="display")
        self.drag_start_index = None
        self.output_directory = str(Path.home())
        self.output_filename = tk.StringVar(value="combined.pdf")
        self.last_output_file = None
        self.preview_window = None
        self.preview_file_index = None
        self.preview_label = None
        
        # Add files button above list
        add_button = tk.Button(
            root,
            text="Add PDFs to Combine...",
            command=self.add_files,
            width=25,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10)
        )
        add_button.pack(padx=10, pady=(10, 5))
        
        # Title above list
        title_label = tk.Label(root, text="Order of Files to Combine:", font=("Arial", 10, "bold"))
        title_label.pack(anchor=tk.W, padx=10, pady=(5, 5))
        
        # Listbox frame
        list_frame = tk.Frame(root)
        list_frame.pack(pady=(0, 10), padx=10, fill=tk.X)
        
        # Column headers using fixed-width labels so columns align with listbox
        header_frame = tk.Frame(list_frame, bg="#E0E0E0")
        header_frame.pack(anchor=tk.W, fill=tk.X)

        hdr_font = ("Courier New", 9)
        # Numbering column header
        num_hdr = tk.Label(header_frame, text="#", font=hdr_font, bg="#E0E0E0", width=4, anchor='e')
        num_hdr.pack(side=tk.LEFT)

        filename_hdr = tk.Label(header_frame, text="Filename", font=hdr_font, bg="#E0E0E0", width=55, anchor='w')
        filename_hdr.pack(side=tk.LEFT)

        size_hdr = tk.Label(header_frame, text="File Size", font=hdr_font, bg="#E0E0E0", width=12, anchor='e')
        size_hdr.pack(side=tk.LEFT)

        date_hdr = tk.Label(header_frame, text="Date", font=hdr_font, bg="#E0E0E0", width=11, anchor='w')
        date_hdr.pack(side=tk.LEFT, padx=(6,0))
        
        # Sub-frame for listbox and scrollbar to keep them aligned
        listbox_scroll_frame = tk.Frame(list_frame)
        listbox_scroll_frame.pack(fill=tk.X)
        
        # Scrollbar
        scrollbar = tk.Scrollbar(listbox_scroll_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Listbox
        self.file_listbox = tk.Listbox(
            listbox_scroll_frame,
            yscrollcommand=scrollbar.set,
            height=10,
            font=("Courier New", 9),
            selectmode=tk.EXTENDED
        )
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)
        
        # Bind drag and drop events
        self.file_listbox.bind("<Button-1>", self.on_mouse_down)
        self.file_listbox.bind("<B1-Motion>", self.on_mouse_drag)
        self.file_listbox.bind("<ButtonRelease-1>", self.on_mouse_up)
        self.file_listbox.bind("<Delete>", lambda e: self.remove_file())
        self.file_listbox.bind("<Motion>", self.on_listbox_hover)
        self.file_listbox.bind("<Leave>", self.on_listbox_leave)
        self.file_listbox.bind("<Double-Button-1>", self.on_file_double_click)
        
        # File count label
        self.count_label = tk.Label(root, text="Files selected: 0", font=("Arial", 9))
        self.count_label.pack(pady=1)
        
        # Drag and drop instruction
        drag_drop_note = tk.Label(
            root,
            text="Single click to select. Ctrl-Click to select multiple. Drag to change order, Hover to see the first page, Double click to open.",
            font=("Arial", 8),
            fg="#666666"
        )
        drag_drop_note.pack(pady=1)
        
        # Sort selection frame (replace previous Combine Order controls)
        sort_frame = tk.LabelFrame(root, text="Combine List Order Sort", font=("Arial", 10, "bold"), padx=10, pady=5)
        sort_frame.pack(pady=5, padx=10, fill=tk.X)

        # Sorting state
        self.sort_key = None  # 'name' | 'size' | 'date'
        self.sort_reverse = False

        # Buttons to sort by Name / Size / Date (click toggles reverse for that key)
        btn_frame = tk.Frame(sort_frame)
        btn_frame.pack(anchor=tk.W)

        self.sort_name_button = tk.Button(btn_frame, text="Name", command=lambda: self.on_sort_clicked('name'), width=12)
        self.sort_name_button.pack(side=tk.LEFT, padx=4)

        self.sort_size_button = tk.Button(btn_frame, text="Size", command=lambda: self.on_sort_clicked('size'), width=12)
        self.sort_size_button.pack(side=tk.LEFT, padx=4)

        self.sort_date_button = tk.Button(btn_frame, text="Date", command=lambda: self.on_sort_clicked('date'), width=12)
        self.sort_date_button.pack(side=tk.LEFT, padx=4)
        
        # Button frame below listbox for Remove/Clear buttons
        listbox_button_frame = tk.Frame(root)
        listbox_button_frame.pack(pady=3)
        
        # Remove selected button
        remove_button = tk.Button(
            listbox_button_frame,
            text="Remove Selected",
            command=self.remove_file,
            width=15,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10)
        )
        remove_button.grid(row=0, column=0, padx=5)
        
        # Clear all button
        clear_button = tk.Button(
            listbox_button_frame,
            text="Clear All",
            command=self.clear_files,
            width=15,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 10)
        )
        clear_button.grid(row=0, column=1, padx=5)
        
        # Output settings frame
        output_frame = tk.LabelFrame(root, text="Output Settings", font=("Arial", 10, "bold"), padx=10, pady=8)
        output_frame.pack(pady=5, padx=10, fill=tk.X)
        
        # Filename frame
        filename_frame = tk.Frame(output_frame)
        filename_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(filename_frame, text="Filename of Combined PDF:", font=("Arial", 9)).pack(side=tk.LEFT, padx=5)
        
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
            highlightbackground="#BBBBBB",
            highlightthickness=1,
            padx=6,
            pady=4,
        )
        dir_box.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        self.location_label = tk.Label(
            dir_box,
            text=self.output_directory,
            font=("Arial", 9, "bold"),
            fg="#000",
            anchor="w"
        )
        self.location_label.pack(side=tk.LEFT)
        # Pack the Browse button to the right of the save-location box
        browse_button.pack(side=tk.LEFT, padx=5)
        
        # Bottom button frame
        bottom_frame = tk.Frame(root)
        bottom_frame.pack(pady=5, padx=10, fill=tk.X)
        
        # Center frame for equal-width buttons
        center_frame = tk.Frame(bottom_frame)
        center_frame.pack()
        
        # Combine button
        combine_button = tk.Button(
            center_frame,
            text="Combine PDFs",
            command=self.combine_pdfs,
            bg="#E0E0E0",
            fg="black",
            font=("Arial", 11),
            height=1,
            width=20
        )
        combine_button.pack(side=tk.LEFT, padx=5)
        
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
            width=20
        )
        quit_button.pack(side=tk.LEFT, padx=5)
    
    def add_files(self):
        """Open file dialog to select PDF files"""
        files = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        added_count = 0
        duplicate_count = 0
        duplicates = []
        
        for file in files:
            if file not in self.pdf_files:
                self.pdf_files.append(file)
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
            self.update_sort_buttons()
        except Exception:
            self.refresh_listbox()

        self.update_count()
        
        # Show warning if duplicates were attempted
        if duplicate_count > 0:
            duplicates_text = "\n".join(f"  • {dup}" for dup in duplicates)
            messagebox.showwarning(
                "Duplicate Files",
                f"The following file(s) are already in the list and were not added:\n\n{duplicates_text}"
            )
    
    def remove_file(self):
        """Remove selected file(s) from list"""
        try:
            selections = self.file_listbox.curselection()
            if not selections:
                messagebox.showwarning("Warning", "Please select a file to remove.")
                return

            # Delete in reverse order to avoid index shifting
            for index in reversed(selections):
                del self.pdf_files[index]

            # Refresh display and count; clear sort state so arrows disappear
            try:
                self.sort_key = None
                self.sort_reverse = False
                self.refresh_listbox()
                self.update_sort_buttons()
            except Exception:
                self.refresh_listbox()

            self.update_count()
        except IndexError:
            messagebox.showwarning("Warning", "Please select a file to remove.")
    
    def clear_files(self):
        """Clear all files from list"""
        self.pdf_files.clear()
        try:
            # Clear any active sort when list is cleared
            self.sort_key = None
            self.sort_reverse = False
            self.refresh_listbox()
            self.update_sort_buttons()
        except Exception:
            self.refresh_listbox()

        self.update_count()
    
    def update_count(self):
        """Update the file count label"""
        count = len(self.pdf_files)
        self.count_label.config(text=f"Files selected: {count}")
    
    def get_file_info(self, file_path: str) -> str:
        """Get formatted file info with three columns: filename, filesize, date"""
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
            # Truncate long filenames so columns remain aligned in the fixed-width listbox
            max_filename_len = 55
            if len(filename) > max_filename_len:
                filename = filename[: max_filename_len - 3] + "..."

            # Format as three columns: filename (55 chars), size (12 chars), date (10 chars)
            return f"{filename:<55} {size_str:>12}  {date_str}"
        except Exception:
            return Path(file_path).name

    def format_list_item(self, index: int, file_path: str) -> str:
        """Return the listbox entry string with numbering plus file info."""
        base = self.get_file_info(file_path)
        return f"{index+1:>3}. {base}"

    def refresh_listbox(self):
        """Rebuild the listbox contents from `self.pdf_files`, applying numbering."""
        self.file_listbox.delete(0, tk.END)
        for i, pdf_file in enumerate(self.pdf_files):
            item = self.format_list_item(i, pdf_file)
            self.file_listbox.insert(tk.END, item)
    
    def on_mouse_down(self, event):
        """Handle mouse down event for drag and drop"""
        self.drag_start_index = self.file_listbox.nearest(event.y)
    
    def on_mouse_drag(self, event):
        """Handle mouse drag event"""
        if self.drag_start_index is None:
            return
        
        current_index = self.file_listbox.nearest(event.y)
        
        if current_index != self.drag_start_index and 0 <= current_index < len(self.pdf_files):
            # Reorder the backing list according to drag-and-drop and refresh display
            dragged_file = self.pdf_files.pop(self.drag_start_index)
            self.pdf_files.insert(current_index, dragged_file)
            self.drag_start_index = current_index

            # Refresh list display and restore selection to the moved item
            self.refresh_listbox()
            try:
                self.file_listbox.selection_set(current_index)
            except Exception:
                pass

            # User manually reordered via drag-and-drop: clear any active sort indicators
            try:
                self.sort_key = None
                self.sort_reverse = False
                self.update_sort_buttons()
            except Exception:
                pass
    
    def on_mouse_up(self, event):
        """Handle mouse up event"""
        self.drag_start_index = None
    
    def on_listbox_hover(self, event):
        """Handle listbox hover to show preview"""
        index = self.file_listbox.nearest(event.y)
        
        if index < 0 or index >= len(self.pdf_files):
            self.hide_preview()
            return
        
        # Only show preview if hovering over a different file
        if self.preview_file_index != index:
            self.show_preview(index, event)
    
    def on_listbox_leave(self, event):
        """Hide preview when mouse leaves listbox"""
        self.hide_preview()
    
    def on_file_double_click(self, event):
        """Open selected PDF file with system default viewer on double-click"""
        selection = self.file_listbox.curselection()
        if not selection:
            return
        
        index = selection[0]
        if 0 <= index < len(self.pdf_files):
            file_path = self.pdf_files[index]
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
    
    def get_pdf_thumbnail(self, file_path: str, width: int = 200, height: int = 280) -> Image.Image:
        """Extract first page of PDF and return as PIL Image"""
        try:
            with open(file_path, 'rb') as pdf_file:
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                
                if len(pdf_reader.pages) == 0:
                    return None
                
                # We'll create a simple text image since PyPDF2 doesn't directly render
                # Extract text from first page to display
                first_page = pdf_reader.pages[0]
                text = first_page.extract_text()[:200]  # First 200 chars
                
                # Create a simple image with the text
                img = Image.new('RGB', (width, height), color='white')
                # Note: We can't easily render text with PIL without additional libraries
                # So we'll just create a placeholder with page count info
                return img
        except Exception:
            return None
    
    def show_preview(self, index: int, event):
        """Show a preview popup with PDF thumbnail"""
        self.hide_preview()
        
        file_path = self.pdf_files[index]
        
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
            x = event.x_root + 15
            y = event.y_root + 15
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
        """Handle sort button clicks. Clicking the same key toggles reverse; clicking a new key sets ascending."""
        if self.sort_key == key:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_key = key
            self.sort_reverse = False

        self.apply_sort()
        self.update_sort_buttons()

    def apply_sort(self):
        """Sort `self.pdf_files` according to current sort_key and sort_reverse, then refresh listbox."""
        try:
            if self.sort_key == 'name':
                self.pdf_files.sort(key=lambda x: Path(x).name.lower(), reverse=self.sort_reverse)
            elif self.sort_key == 'size':
                def _size_key(p):
                    try:
                        return os.path.getsize(p)
                    except Exception:
                        return -1
                self.pdf_files.sort(key=_size_key, reverse=self.sort_reverse)
            elif self.sort_key == 'date':
                def _date_key(p):
                    try:
                        return os.path.getmtime(p)
                    except Exception:
                        return 0
                self.pdf_files.sort(key=_date_key, reverse=self.sort_reverse)
            # If sort_key is None, do nothing (preserve display order)

            # Update listbox (with numbering)
            self.refresh_listbox()
        except Exception:
            pass

    def update_sort_buttons(self):
        """Update button labels to show sort direction for the active key."""
        up = '▲'
        down = '▼'

        # Reset labels
        self.sort_name_button.config(text='Name')
        self.sort_size_button.config(text='Size')
        self.sort_date_button.config(text='Date')

        if self.sort_key == 'name':
            arrow = down if self.sort_reverse else up
            self.sort_name_button.config(text=f'Name {arrow}')
        elif self.sort_key == 'size':
            arrow = down if self.sort_reverse else up
            self.sort_size_button.config(text=f'Size {arrow}')
        elif self.sort_key == 'date':
            arrow = down if self.sort_reverse else up
            self.sort_date_button.config(text=f'Date {arrow}')
    
    def browse_output_location(self):
        """Open directory browser to select output location"""
        directory = filedialog.askdirectory(
            title="Select output location",
            initialdir=self.output_directory
        )
        
        if directory:
            self.output_directory = directory
            self.location_label.config(text=self.output_directory)
    
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

        # Determine the order of PDFs to combine
        files_to_combine = self.pdf_files.copy()
        
        if self.combine_order.get() == "alphabetical":
            # Sort by filename alphabetically
            files_to_combine.sort(key=lambda x: Path(x).name.lower())
        
        # Show summary before combining
        self.show_combine_summary(output_file, files_to_combine)
    
    def show_combine_summary(self, output_file, files_to_combine):
        """Show a summary of PDFs to combine before proceeding"""
        # Calculate total pages and file size
        total_pages = 0
        total_size_bytes = 0
        
        try:
            for pdf_file in files_to_combine:
                # Count pages
                with open(pdf_file, 'rb') as f:
                    pdf_reader = PyPDF2.PdfReader(f)
                    total_pages += len(pdf_reader.pages)
                
                # Get file size
                total_size_bytes += os.path.getsize(pdf_file)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read PDF information: {e}")
            return
        
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
        summary_window.geometry("400x320")
        summary_window.resizable(False, False)
        summary_window.transient(self.root)
        summary_window.grab_set()
        
        # Center the summary window
        summary_window.update_idletasks()
        x = (summary_window.winfo_screenwidth() // 2) - (400 // 2)
        y = (summary_window.winfo_screenheight() // 2) - (320 // 2)
        summary_window.geometry(f"400x320+{x}+{y}")
        
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
        files_label = tk.Label(
            info_frame,
            text=f"Files to combine:  {len(files_to_combine)} files",
            font=("Arial", 10, "bold"),
            fg="#0066CC",
            anchor="w"
        )
        files_label.pack(fill=tk.X, pady=5)
        
        # Total pages
        pages_label = tk.Label(
            info_frame,
            text=f"Total pages:  {total_pages} pages",
            font=("Arial", 10, "bold"),
            fg="#0066CC",
            anchor="w"
        )
        pages_label.pack(fill=tk.X, pady=5)
        
        # Total size
        size_label = tk.Label(
            info_frame,
            text=f"Total size:  {size_str}",
            font=("Arial", 10, "bold"),
            fg="#0066CC",
            anchor="w"
        )
        size_label.pack(fill=tk.X, pady=5)
        
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
            pdf_merger = None
            try:
                # Create PDF merger object
                pdf_merger = PyPDF2.PdfMerger()
                
                # Add all PDFs to merger in the selected order
                for i, pdf_file in enumerate(files_to_combine):
                    # Check if cancelled
                    if cancel_flag['cancelled']:
                        self.root.after(0, lambda: (
                            progress_window.destroy(),
                            messagebox.showinfo("Cancelled", "PDF combining operation was cancelled.")
                        ))
                        if pdf_merger:
                            pdf_merger.close()
                        return
                    
                    # Update progress
                    self.root.after(0, lambda idx=i, f=pdf_file: (
                        progress_label.config(text=f"Processing: {Path(f).name}"),
                        progress_bar.config(value=idx),
                        counter_label.config(text=f"{idx} of {len(files_to_combine)} files processed")
                    ))
                    
                    pdf_merger.append(pdf_file)
                
                # Check if cancelled before writing
                if cancel_flag['cancelled']:
                    self.root.after(0, lambda: (
                        progress_window.destroy(),
                        messagebox.showinfo("Cancelled", "PDF combining operation was cancelled.")
                    ))
                    if pdf_merger:
                        pdf_merger.close()
                    return
                
                # Update for writing phase
                self.root.after(0, lambda: (
                    progress_label.config(text="Writing combined PDF..."),
                    progress_bar.config(value=len(files_to_combine)),
                    counter_label.config(text=f"{len(files_to_combine)} of {len(files_to_combine)} files processed"),
                    cancel_button.config(state='disabled')
                ))
                
                # Write combined PDF
                pdf_merger.write(output_file)
                pdf_merger.close()
                
                # Remember the output file
                self.last_output_file = output_file
                
                # Close progress window and show success
                self.root.after(0, lambda: (
                    progress_window.destroy(),
                    self.show_success_dialog(output_file)
                ))
                
            except FileNotFoundError as e:
                if pdf_merger:
                    pdf_merger.close()
                self.root.after(0, lambda: (
                    progress_window.destroy(),
                    messagebox.showerror("Error", f"File not found: {e}")
                ))
            except Exception as e:
                if pdf_merger:
                    pdf_merger.close()
                self.root.after(0, lambda: (
                    progress_window.destroy(),
                    messagebox.showerror("Error", f"An error occurred: {str(e)}")
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
