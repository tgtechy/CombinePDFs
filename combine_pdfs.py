from __future__ import annotations
print(">>> RUNNING combine_pdfs.py FROM:", __file__)

# ---------------------------------------------------------------------------
# Reusable helpers for widget state and custom dialogs
# ---------------------------------------------------------------------------
import tkinter as tk
from tkinter import ttk
from typing import List, Optional

def set_widgets_state(widgets: List[tk.Widget], enabled: bool):
    """Enable or disable a list of Tkinter widgets."""
    state = "normal" if enabled else "disabled"
    for w in widgets:
        try:
            w.config(state=state)
        except Exception:
            pass

def show_custom_dialog(
    parent,
    title: str,
    message: str,
    icon: Optional[str] = None,  # path to image or emoji
    buttons: List[str] = ["OK"],
    default: Optional[str] = None,
    cancel: Optional[str] = None,
    width: int = 420,
    height: int = 180,
    wraplength: int = 320,
    style: Optional[str] = None,
) -> str:
    """Show a custom dialog with message, icon, and custom buttons. Returns the label of the button pressed."""
    dlg = tk.Toplevel(parent)
    dlg.title(title)
    dlg.transient(parent)
    dlg.grab_set()
    dlg.resizable(False, False)
    # Set icon if available
    try:
        from pathlib import Path
        icon_path = Path(__file__).resolve().parent / "pdfcombinericon.ico"
        if icon_path.exists():
            dlg.iconbitmap(str(icon_path))
    except Exception:
        pass
    # Center dialog
    parent.update_idletasks()
    px = parent.winfo_rootx()
    py = parent.winfo_rooty()
    pw = parent.winfo_width()
    ph = parent.winfo_height()
    x = px + (pw // 2) - (width // 2)
    y = py + (ph // 2) - (height // 2)
    dlg.geometry(f"{width}x{height}+{x}+{y}")
    bg = parent.cget('background')
    dlg.configure(bg=bg)
    dlg.columnconfigure(1, weight=1)
    # Icon
    icon_label = None
    if icon and isinstance(icon, str):
        if icon.endswith(('.png', '.gif')):
            try:
                img = tk.PhotoImage(file=icon)
                icon_label = tk.Label(dlg, image=img, bg=bg)
                icon_label.image = img
            except Exception:
                icon_label = tk.Label(dlg, text="!", font=("Segoe UI", 24, "bold"), fg="red", bg=bg)
        else:
            icon_label = tk.Label(dlg, text=icon, font=("Segoe UI", 24, "bold"), fg="red", bg=bg)
    if not icon_label:
        icon_label = tk.Label(dlg, text="!", font=("Segoe UI", 24, "bold"), fg="red", bg=bg)
    icon_label.grid(row=0, column=0, rowspan=2, padx=(24, 12), pady=(24, 12), sticky="n")
    # Message
    msg_label = tk.Label(
        dlg,
        text=message,
        font=("Segoe UI", 11),
        bg=bg,
        anchor="w",
        justify="left",
        wraplength=wraplength
    )
    msg_label.grid(row=0, column=1, columnspan=2, sticky="w", padx=(0, 16), pady=(24, 24))
    # Buttons
    result = [None]
    def on_btn(label):
        result[0] = label
        dlg.destroy()
    btn_frame = ttk.Frame(dlg)
    btn_frame.grid(row=1, column=0, columnspan=3, pady=(0, 24))
    for i, label in enumerate(buttons):
        btn = ttk.Button(btn_frame, text=label, command=lambda l=label: on_btn(l), style=style or "WinButton.TButton", width=14)
        btn.grid(row=0, column=i, padx=(0 if i == 0 else 12, 0))
        if default and label == default:
            btn.focus_set()
        if cancel and label == cancel:
            dlg.protocol("WM_DELETE_WINDOW", lambda: on_btn(cancel))
    if not cancel:
        dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)
    dlg.wait_window()
    return result[0]

import os
import sys
import threading
from pathlib import Path
from dataclasses import dataclass
from typing import List, Optional
 # Removed TKinterModernThemes integration

import ttkbootstrap as tb
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

__VERSION__ = "2.0.0"

# ---------------------------------------------------------------------------
# Define ROOT and CORE_DIR FIRST
# ---------------------------------------------------------------------------

ROOT = Path(__file__).resolve().parent
CORE_DIR = ROOT / "core"

# ---------------------------------------------------------------------------
# Normalize paths to prevent UNC canonicalization
# ---------------------------------------------------------------------------

ROOT = Path(os.path.abspath(str(ROOT)))
CORE_DIR = Path(os.path.abspath(str(CORE_DIR)))

# ---------------------------------------------------------------------------
# Clean sys.path
# ---------------------------------------------------------------------------

# Remove UNC paths
sys.path = [p for p in sys.path if not p.startswith("\\\\")]

# Force normalized local paths
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(CORE_DIR))

# ---------------------------------------------------------------------------
# Ensure project root is on sys.path
# ---------------------------------------------------------------------------

if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))
    
# ---------------------------------------------------------------------------
# Core imports
# ---------------------------------------------------------------------------
from core.settings import load_settings, save_settings
from core.file_manager import (
    add_files_to_list,
    move_up,
    move_down,
    remove_file,
    clear_files,
)

print("\n=== MODULE ORIGIN DIAGNOSTIC ===")
import core
print("core module:", core.__file__)
import core.settings
print("settings module:", core.settings.__file__)
import core.page_ops
print("page_ops module:", core.page_ops.__file__)
import core.image_tools
print("image_tools module:", core.image_tools.__file__)
import core.compression
print("compression module:", core.compression.__file__)
import core.toc
print("toc module:", core.toc.__file__)
print("=== END MODULE ORIGIN ===\n")
from core.pdf_merger import merge_files, FileEntry, MergeOptions


import platform
# ---------------------------------------------------------------------------
# App settings model
# ---------------------------------------------------------------------------

def get_user_config_path():
    if platform.system() == "Windows":
        appdata = os.environ.get("APPDATA")
        if appdata:
            return Path(appdata) / "CombinePDFs" / "settings.json"
    # Fallback for other OS
    home = Path.home()
    return home / ".combinepdfs_settings.json"

CONFIG_PATH = get_user_config_path()


@dataclass
class AppSettings:
    last_open_dir: str = ""
    last_save_dir: str = ""
    output_filename: str = ""

    # General
    delete_blank_pages: bool = False
    insert_toc: bool = False
    breaker_uniform_size: bool = False
    add_filename_bookmarks: bool = False
    add_breaker_pages: bool = False

    # Compression
    compression_enabled: bool = False
    compression_level: str = "Medium"

    # Watermark
    watermark_enabled: bool = False
    watermark_text: str = ""
    watermark_opacity: float = 0.3
    watermark_rotation: int = 0
    watermark_position: str = "Center"
    watermark_font_size: int = 24
    watermark_safe_mode: bool = True
    watermark_font_color: str = "#000000"

    # Metadata
    metadata_enabled: bool = False
    pdf_title: str = ""
    pdf_author: str = ""
    pdf_subject: str = ""
    pdf_keywords: str = ""

    # Scaling
    scaling_enabled: bool = False
    scaling_mode: str = "Fit"
    scaling_percent: int = 100

    # UI
    dark_mode: bool = False

    @classmethod
    def from_dict(cls, data: dict) -> "AppSettings":
        # Only use keys that are defined in the class
        valid_keys = set(cls().__dict__.keys())
        filtered = {k: v for k, v in data.items() if k in valid_keys}
        return cls(**{**cls().__dict__, **filtered})

    def to_dict(self) -> dict:
        return self.__dict__.copy()


# ---------------------------------------------------------------------------
# Progress dialog
# ---------------------------------------------------------------------------

class ProgressDialog:
    def __init__(self, parent: tk.Tk, title: str = "Processing...") -> None:
        self.parent = parent
        self.cancelled = False

        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.transient(parent)
        self.top.grab_set()
        self.top.resizable(False, False)
        # Set icon for dialog window
        icon_path = Path(__file__).resolve().parent / "pdfcombinericon.ico"
        if icon_path.exists():
            try:
                self.top.iconbitmap(str(icon_path))
            except Exception:
                pass

        # Set a larger size
        width, height = 380, 200
        self.top.geometry(f"{width}x{height}")

        # Center the dialog in the parent window
        parent.update_idletasks()
        root_x = parent.winfo_rootx()
        root_y = parent.winfo_rooty()
        root_w = parent.winfo_width()
        root_h = parent.winfo_height()
        x = root_x + (root_w // 2) - (width // 2)
        y = root_y + (root_h // 2) - (height // 2)
        self.top.geometry(f"{width}x{height}+{x}+{y}")

        self.top.columnconfigure(0, weight=1)
        bg = self.top.cget("background")
        ttk.Label(self.top, text="Processing, please wait...", background=bg).grid(row=0, column=0, padx=10, pady=(16, 2), sticky="nsew")
        ttk.Label(self.top, text="Merging:", background=bg).grid(row=1, column=0, padx=10, pady=(2, 2), sticky="nsew")
        self.filename_var = tk.StringVar(value="")
        self.filename_label = ttk.Label(self.top, textvariable=self.filename_var, font=("Segoe UI", 10, "bold"), background=bg)
        self.filename_label.grid(row=2, column=0, padx=10, pady=(0, 8), sticky="nsew")
        self.progress = ttk.Progressbar(self.top, mode="indeterminate", length=250)
        self.progress.grid(row=3, column=0, padx=10, pady=5, sticky="nsew")
        self.progress.start(10)

        from ttkbootstrap import Button as TBButton
        TBButton(self.top, text="Cancel", command=self._on_cancel, style="WinButton.TButton").grid(row=4, column=0, padx=10, pady=(10, 18), sticky="nsew")

        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def set_filename(self, filename: str):
        self.filename_var.set(filename)

    def _on_cancel(self) -> None:
        self.cancelled = True

    def close(self) -> None:
        self.progress.stop()
        self.top.grab_release()
        self.top.destroy()


# ---------------------------------------------------------------------------
# Main UI class
# ---------------------------------------------------------------------------

class CombinePDFsUI:
    
    def _get_selected_index(self):
        """Return the first selected index in the file list as int, or None if nothing is selected."""
        sels = self.tree.selection()
        if not sels:
            return None
        # Treeview item IDs are string indices
        try:
            return int(sels[0])
        except Exception:
            return None
    def on_browse_output(self) -> None:
        from tkinter import filedialog
        import os
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title="Select output PDF file",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=getattr(self.settings, 'last_save_dir', os.getcwd()),
            initialfile="Combined.pdf"
        )
        if not path:
            return
        # Ensure .pdf extension
        if not path.lower().endswith('.pdf'):
            path += '.pdf'
        self.output_var.set(path)
        if hasattr(self.settings, 'last_save_dir'):
            self.settings.last_save_dir = os.path.dirname(path)
        self.settings.output_filename = path
        self._save_app_settings()
    
    def _setup_styles(self):
        # Patch ttkbootstrap palette to force all buttons to use light gray background and black text
        style = tb.Style()
        # The following palette patch is disabled due to incompatibility with recent ttkbootstrap versions
        # which do not support item assignment on the Colors object.
        # If you want to customize button colors, use style.configure or style.map below instead.
        # Force base TButton style to be light gray with black text (overrides ttkbootstrap dark theme)
        style.configure('TButton',
            background='#E1E1E1',
            foreground='black',
            font=('Segoe UI', 9, 'normal'),
            bordercolor='#A9A9A9',
            borderwidth=1,
            focusthickness=2,
            focuscolor='#A9A9A9',
            relief='raised',
            padding=(6, 2),
        )
        style.map('TButton',
            background=[('active', '#D5D5D5'), ('pressed', '#C8C8C8'), ('!active', '#E1E1E1'), ('disabled', '#E1E1E1')],
            foreground=[('active', 'black'), ('pressed', 'black'), ('!active', 'black'), ('disabled', '#888888')],
            bordercolor=[('focus', '#A9A9A9'), ('!focus', '#A9A9A9')],
        )
        # Remove all manual theme sourcing and set_theme calls
        style.configure('Custom.TNotebook',
            borderwidth=2,
            relief='groove',
            background='#f7f7f7',
        )
        style.configure('Custom.TNotebook.Tab',
            padding=[18, 6],
            background="#f7f7f7",
            borderwidth=1,
            relief='flat',
            foreground='black',
            font=('Segoe UI', 9, 'normal'),
            lightcolor='#e1e1e1',
            bordercolor='#a9a9a9',
        )
        style.map('Custom.TNotebook.Tab',
            background=[('selected', "#ffffff"), ('active', "#e1e1e1"), ('!selected', "#f7f7f7")],
            foreground=[('selected', 'black'), ('active', 'black'), ('!selected', '#888888')],
            bordercolor=[('selected', '#a9a9a9'), ('active', '#a9a9a9'), ('!selected', '#e1e1e1')],
            font=[('selected', ('Segoe UI', 9, 'normal')), ('active', ('Segoe UI', 9, 'normal')), ('!selected', ('Segoe UI', 9, 'normal'))],
            padding=[('selected', [18, 6]), ('active', [18, 6]), ('!selected', [18, 6])],
        )

        # Set ttk.Checkbutton background to match window background
        window_bg = self.root.cget('background') if hasattr(self, 'root') else '#f0f0f0'
        style.configure('TCheckbutton', background=window_bg)

        # Custom Windows 10/11 light-mode button style
        style.configure('WinButton.TButton',
            background='#E1E1E1',
            foreground='black',
            font=('Segoe UI', 9, 'normal'),
            bordercolor='#A9A9A9',
            borderwidth=1,
            focusthickness=2,
            focuscolor='#A9A9A9',
            relief='raised',
            padding=(6, 2),
        )
        style.map('WinButton.TButton',
            background=[('active', '#D5D5D5'), ('pressed', '#C8C8C8'), ('!active', '#E1E1E1'), ('disabled', '#E1E1E1')],
            foreground=[('active', 'black'), ('pressed', 'black'), ('!active', 'black'), ('disabled', '#888888')],
            bordercolor=[('focus', '#A9A9A9'), ('!focus', '#A9A9A9')],
        )

        # Restyle buttons for light mode to match Windows standard and reduce size
        # Only apply if not in dark mode
        if not getattr(self.settings, 'dark_mode', False):
            # Custom style for all buttons in light mode
            style.configure('CustomLight.TButton',
                background='#f0f0f0',
                foreground='black',
                font=('Segoe UI', 9, 'normal'),
                borderwidth=1,
                focusthickness=2,
                focuscolor='SystemWindowFrame',
                relief='raised',
                padding=(2, 0),
            )
            style.map('CustomLight.TButton',
                background=[('active', '#e5e5e5'), ('pressed', '#e5e5e5'), ('!active', '#f0f0f0')],
                foreground=[('active', 'black'), ('pressed', 'black'), ('!active', 'black')],
                font=[('active', ('Segoe UI', 9, 'normal')), ('!active', ('Segoe UI', 9, 'normal'))],
            )

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Combine PDFs")

        self.files: List[dict] = []
        self.settings = self._load_app_settings()

        self._build_ui()

        self._merge_thread: Optional[threading.Thread] = None
        self._progress_dialog: Optional[ProgressDialog] = None

        # Save settings on exit
        self.root.protocol("WM_DELETE_WINDOW", self._on_exit)

    def _on_exit(self):
        # Sync UI values to settings before saving
        try:
            self.settings.delete_blank_pages = self.var_delete_blank.get()
            self.settings.insert_toc = self.var_insert_toc.get()
            self.settings.breaker_uniform_size = self.var_breaker_uniform.get()
            self.settings.add_filename_bookmarks = self.var_add_filename_bookmarks.get()
            self.settings.compression_enabled = self.var_comp_enabled.get()
            self.settings.compression_level = self.var_comp_level.get()
            self.settings.watermark_enabled = self.var_wm_enabled.get()
            self.settings.watermark_text = self.var_wm_text.get()
            self.settings.watermark_opacity = float(self.var_wm_opacity.get())
            self.settings.watermark_rotation = int(self.var_wm_rotation.get())
            self.settings.watermark_position = self.var_wm_position.get()
            self.settings.watermark_font_size = int(self.var_wm_font_size.get())
            self.settings.watermark_safe_mode = self.var_wm_safe.get()
            self.settings.watermark_font_color = self.var_wm_font_color.get()
            self.settings.metadata_enabled = self.var_meta_enabled.get()
            self.settings.pdf_title = self.var_meta_title.get()
            self.settings.pdf_author = self.var_meta_author.get()
            self.settings.pdf_subject = self.var_meta_subject.get()
            self.settings.pdf_keywords = self.var_meta_keywords.get()
            self.settings.scaling_enabled = self.var_scale_enabled.get()
            self.settings.scaling_mode = self.var_scale_mode.get()
            self.settings.scaling_percent = int(self.var_scale_percent.get())
            self.settings.add_breaker_pages = self.var_add_breaker_pages.get()
            self.settings.dark_mode = self.var_dark_mode.get()
        except Exception:
            pass
        self._save_app_settings()
        self.root.destroy()

    # -----------------------------------------------------------------------
    # Settings persistence
    # -----------------------------------------------------------------------

    def _load_app_settings(self) -> AppSettings:
        data = load_settings(CONFIG_PATH)
        settings = AppSettings.from_dict(data)
        # Ensure font color is always set
        if not getattr(settings, 'watermark_font_color', None):
            settings.watermark_font_color = '#000000'
        return settings

    def _save_app_settings(self) -> None:
        # Ensure font color is saved
        if not getattr(self.settings, 'watermark_font_color', None):
            self.settings.watermark_font_color = '#000000'
        save_settings(CONFIG_PATH, self.settings.to_dict())

    # -----------------------------------------------------------------------
    # UI construction
    # -----------------------------------------------------------------------

    def _build_ui(self) -> None:
        self.status_var = tk.StringVar()
        self._update_status_bar()
        self._setup_styles()
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)


        cb_bg = "#dcdad5"
        main = tk.Frame(self.root, bg=cb_bg)
        main.grid(row=0, column=0, sticky="nsew")
        main.rowconfigure(0, weight=1)
        main.columnconfigure(0, weight=1)

        # Notebook for two main tabs with custom style
        notebook = ttk.Notebook(main, style='Custom.TNotebook')
        notebook.grid(row=0, column=0, sticky="nsew", padx=(10,10), pady=(16, 16))
        main.rowconfigure(0, weight=1)
        main.columnconfigure(0, weight=1)

        # --- Files Tab ---
        files_tab = tk.Frame(notebook, bg=cb_bg)
        files_tab.rowconfigure(0, weight=1)
        files_tab.columnconfigure(0, weight=1)
        self._build_file_list(files_tab)
        self._build_file_buttons(files_tab)
        notebook.add(files_tab, text="📄 File List")

        # --- Options Tab ---

        options_tab = tk.Frame(notebook, bg=cb_bg)
        options_tab.rowconfigure(0, weight=1)
        options_tab.columnconfigure(0, weight=1)
        self._build_settings_notebook(options_tab)
        notebook.add(options_tab, text="⚙️ Options & Settings")

        # --- Instructions Tab ---
        instructions_tab = tk.Frame(notebook, bg=cb_bg)
        instructions_tab.rowconfigure(0, weight=1)
        instructions_tab.columnconfigure(0, weight=1)

        # Add a scrollable Text widget
        from tkinter import scrolledtext
        instructions_textbox = scrolledtext.ScrolledText(
            instructions_tab,
            wrap="word",
            font=("Segoe UI", 10),
            bg=cb_bg,
            relief="flat",
            borderwidth=0,
            state="normal",
            height=30,
            width=90
        )
        instructions_textbox.grid(row=0, column=0, sticky="nsew", padx=16, pady=16)
        instructions_tab.rowconfigure(0, weight=1)
        instructions_tab.columnconfigure(0, weight=1)

        # Load instructions.md content
        try:
            with open(os.path.join(ROOT, "instructions.md"), "r", encoding="utf-8") as f:
                instructions_content = f.read()
        except Exception as e:
            instructions_content = f"Could not load instructions.md: {e}"
        instructions_textbox.insert("1.0", instructions_content)
        instructions_textbox.config(state="disabled")
        notebook.add(instructions_tab, text="❓ Instructions")

        # --- Output controls at the bottom ---
        bottom = ttk.Frame(self.root, padding=(10, 0, 10, 10))
        bottom.grid(row=1, column=0, sticky="ew")
        bottom.columnconfigure(1, weight=1)

        ttk.Label(bottom, text="Output file:").grid(row=0, column=0, sticky="w")
        # Initialize output_var from settings and persist changes
        self.output_var = tk.StringVar(value=getattr(self.settings, 'output_filename', ''))
        output_entry = ttk.Entry(bottom, textvariable=self.output_var)
        output_entry.grid(row=0, column=1, sticky="ew", padx=4)
        def on_output_var_change(*args):
            self.settings.output_filename = self.output_var.get()
            self._save_app_settings()
        self.output_var.trace_add('write', on_output_var_change)
        # Use bootstyle for Windows-like button in light mode
        from ttkbootstrap import Button as TBButton
        TBButton(bottom, text="Browse...", command=self.on_browse_output, style="WinButton.TButton").grid(row=0, column=2)
        # Remove Merge button from here


        # Merge, Quit, and Donate link row
        bottom_frame = ttk.Frame(self.root, padding=(10, 0, 10, 20))
        bottom_frame.grid(row=2, column=0, sticky="ew", pady=0)
        # Use pack: donate link left, button group centered with expand
        btn_width = 16


        # Donate link at the very bottom, centered
        def open_donate_link(event=None):
            import webbrowser
            webbrowser.open_new("https://paypal.me/tgtechdevshop")

        def open_github_link(event=None):
            import webbrowser
            webbrowser.open_new("https://github.com/tgtechy/CombinePDFs")


        # Centered button group
        btns_frame = ttk.Frame(bottom_frame)
        btns_frame.pack(side="left", expand=True, anchor="center")
        TBButton(btns_frame, text="Merge", command=self.on_merge_clicked, style="WinButton.TButton", width=btn_width).pack(side="left", padx=(0, 10))
        TBButton(btns_frame, text="Quit", command=self._on_exit, style="WinButton.TButton", width=btn_width).pack(side="left", padx=(10, 0))


        # Donate link and version above the status bar, same row
        link_version_frame = ttk.Frame(self.root)
        link_version_frame.grid(row=3, column=0, sticky="ew", pady=0)
        link_version_frame.columnconfigure(0, weight=1)
        link_version_frame.columnconfigure(1, weight=1)

        donate_label = tk.Label(
            link_version_frame,
            text="Like this? Donate!",
            fg="#1976d2",
            cursor="hand2",
            font=("Segoe UI", 8, "underline")
        )
        donate_label.grid(row=0, column=0, sticky="w", padx=4)
        donate_label.bind("<Button-1>", open_donate_link)
        # Force blue color in case theme overrides it
        donate_label.config(fg="#1976d2", foreground="#1976d2", highlightcolor="#1976d2", activeforeground="#1976d2")

        version_label = tk.Label(
            link_version_frame,
            text=f"©2026 tgtech v{__VERSION__}",
            fg="#1976d2",
            cursor="hand2",
            font=("Segoe UI", 8, "underline")
        )
        version_label.grid(row=0, column=1, sticky="e", padx=(0, 10))
        version_label.bind("<Button-1>", open_github_link)
        version_label.config(fg="#1976d2", foreground="#1976d2", highlightcolor="#1976d2", activeforeground="#1976d2")

        # Status bar
        self.status_var = tk.StringVar()
        self._update_status_bar()
        status_bar = ttk.Label(self.root, textvariable=self.status_var, anchor="w", relief="sunken", padding=(8, 2))
        status_bar.grid(row=4, column=0, sticky="ew")

    def _update_status_bar(self):
        count = len(self.files)
        total_size = 0
        for entry in self.files:
            path = entry.get("path")
            try:
                total_size += os.path.getsize(path)
            except Exception:
                pass
        def human_size(size):
            if size is None:
                return "Unknown size"
            for unit in ['bytes', 'KB', 'MB', 'GB', 'TB']:
                if size < 1024.0:
                    return f"{size:.1f} {unit}" if unit != 'bytes' else f"{size:,} bytes"
                size /= 1024.0
            return f"{size:.1f} PB"
        size_str = human_size(total_size)
        self.status_var.set(f"Files: {count}     Total Size: {size_str}")

    # Update status bar after file list changes
    def _build_file_list(self, parent: ttk.Frame) -> None:
        # --- File preview popup ---
        self._preview_popup = None
        self._preview_popup_img = None

        filelist_frame = ttk.Frame(parent)
        filelist_frame.grid(row=0, column=0, sticky="nsew")
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)

        self._filelist_frame = filelist_frame  # Store for instructional label

        self.tree = ttk.Treeview(
            filelist_frame,
            columns=("path", "size", "date", "page_range", "rotation", "reverse"),
            show="headings",
            height=7,
        )
        self._tree_sort_column = None
        self._tree_sort_reverse = False
        self._tree_column_titles = {
            "path": "File",
            "size": "Size",
            "date": "Date",
            "page_range": "Pages",
            "rotation": "Rot",
            "reverse": "Rev",
        }
        self._set_tree_headings()
        # Set column widths and stretch
        # Dynamically size columns to fill the width of the file list box
        total_width = 800  # Approximate initial width of the file list area
        col_widths = {
            "path": int(total_width * 0.55),      # File name gets the most space
            "size": int(total_width * 0.08),      # Size
            "date": int(total_width * 0.1),      # Date Modified
            "page_range": int(total_width * 0.08),# Pages
            "rotation": int(total_width * 0.04),  # Rotation
            "reverse": int(total_width * 0.04),   # Reverse
        }
        for col, width in col_widths.items():
            self.tree.column(col, width=width, anchor=("w" if col=="path" else "center"), stretch=True)

        self.tree.grid(row=0, column=0, sticky="nsew")
        # Make columns auto-stretch when resizing
        for i, col in enumerate(self.tree["columns"]):
            self.tree.column(col, stretch=True)
        vsb = ttk.Scrollbar(filelist_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")
        hsb = ttk.Scrollbar(filelist_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=hsb.set)
        hsb.grid(row=1, column=0, sticky="ew")
        filelist_frame.rowconfigure(0, weight=1)
        filelist_frame.columnconfigure(0, weight=1)
        # Bind double-click to edit cell
        self.tree.bind('<Double-1>', self._on_tree_double_click)
        # Bind preview events
        self._preview_after_id = None
        self.tree.bind('<Motion>', self._on_tree_motion)
        self.tree.bind('<Leave>', self.hide_preview)
        # --- End file list subframe ---
        self._refresh_tree()
        self._update_status_bar()
        return

    def _on_tree_motion(self, event):
        if self._preview_after_id:
            self.tree.after_cancel(self._preview_after_id)
        self._preview_after_id = self.tree.after(400, lambda: self.show_preview(event))

    def show_preview(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid:
            self.hide_preview()
            return
        idx = int(iid)
        entry = self.files[idx]
        path = entry["path"]
        ext = os.path.splitext(path)[1].lower()
        import PIL.Image, PIL.ImageTk
        preview_img = None
        preview_text = None
        if ext in ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.tif'):
            try:
                img = PIL.Image.open(path)
                img.thumbnail((240, 240))
                preview_img = PIL.ImageTk.PhotoImage(img)
            except Exception as e:
                preview_text = f"Image preview failed: {e}"
        elif ext == '.pdf':
            try:
                import fitz
                doc = fitz.open(path)
                page = doc.load_page(0)
                pix = page.get_pixmap(matrix=fitz.Matrix(2,2))
                img = PIL.Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                img.thumbnail((240, 240))
                preview_img = PIL.ImageTk.PhotoImage(img)
            except Exception as e:
                preview_text = f"PDF preview failed: {e}"
        else:
            preview_text = "No preview available."
        if self._preview_popup:
            self._preview_popup.destroy()
            self._preview_popup = None
        self._preview_popup = tk.Toplevel(self.tree)
        self._preview_popup.overrideredirect(True)
        self._preview_popup.attributes("-topmost", True)
        x = self.tree.winfo_pointerx()
        y = self.tree.winfo_pointery()
        self._preview_popup.geometry(f"+{x+16}+{y+16}")
        # Prepare full path and filename label for the bottom
        preview_width = 240
        import textwrap
        chars_per_line = max(22, int(preview_width / 7))
        wrapped_path = '\n'.join(textwrap.wrap(path, chars_per_line))
        info_label = tk.Label(
            self._preview_popup,
            text=wrapped_path,
            bg="white",
            fg="black",
            font=("Segoe UI", 9, "normal"),
            anchor="w",
            justify="left"
        )
        # Place preview image or text first
        if preview_img:
            self._preview_popup_img = preview_img
            label = tk.Label(self._preview_popup, image=preview_img, bg="white", bd=2, relief="solid")
            label.pack()
        else:
            label = tk.Label(self._preview_popup, text=preview_text or "No preview", bg="white", bd=2, relief="solid")
            label.pack()
        # Now pack the info_label at the bottom
        info_label.pack(fill="x", padx=4, pady=(4,2), side="bottom")

    def hide_preview(self, event=None):
        if self._preview_after_id:
            self.tree.after_cancel(self._preview_after_id)
            self._preview_after_id = None
        if self._preview_popup:
            self._preview_popup.destroy()
            self._preview_popup = None
            self._preview_popup_img = None

    def _set_tree_headings(self):
        # Add sort indicator to sorted column
        for col in ("path", "size", "date", "page_range", "rotation", "reverse"):
            title = self._tree_column_titles[col]
            if col == self._tree_sort_column:
                arrow = "▼" if self._tree_sort_reverse else "▲"
                title = f"{title} {arrow}"
            if col in ("path", "size", "date"):
                self.tree.heading(col, text=title, command=lambda c=col: self._sort_tree_column(c))
            else:
                self.tree.heading(col, text=title)

    def _sort_tree_column(self, col):
        # Toggle sort order if same column, else ascending
        if self._tree_sort_column == col:
            self._tree_sort_reverse = not self._tree_sort_reverse
        else:
            self._tree_sort_column = col
            self._tree_sort_reverse = False

        # Prepare sort key
        def get_sort_key(entry):
            path = entry["path"]
            if col == "path":
                return os.path.basename(path).lower()
            elif col == "size":
                try:
                    return os.stat(path).st_size
                except Exception:
                    return -1
            elif col == "date":
                try:
                    return os.stat(path).st_mtime
                except Exception:
                    return 0
            return None

        self.files.sort(key=get_sort_key, reverse=self._tree_sort_reverse)
        self._set_tree_headings()
        self._refresh_tree()

    def _on_tree_double_click(self, event):
        # Identify row and column
        region = self.tree.identify('region', event.x, event.y)
        if region != 'cell':
            return
        row_id = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        col_index = int(col.replace('#', '')) - 1
        if not row_id:
            return
        idx = int(row_id)
        entry = self.files[idx]

        bbox = self.tree.bbox(row_id, col)
        if not bbox:
            return
        x, y, width, height = bbox
        abs_x = self.tree.winfo_rootx() + x
        abs_y = self.tree.winfo_rooty() + y

        # Page Range (Entry)
        # Columns: 0=path, 1=size, 2=date, 3=page_range, 4=rotation, 5=reverse
        if col_index == 3:
            top = tk.Toplevel(self.tree)
            top.overrideredirect(True)
            top.geometry(f"{width}x{height}+{abs_x}+{abs_y}")
            # If no page_range is set or is 'All', show blank; else retain existing value
            current_range = entry.get("page_range", "")
            if not current_range or current_range.strip().lower() == "all":
                var = tk.StringVar(value="")
            else:
                var = tk.StringVar(value=current_range)
            entry_widget = tk.Entry(top, textvariable=var)
            entry_widget.pack(fill="both", expand=True)
            entry_widget.focus_set()

            def on_commit(event=None):
                entry["page_range"] = var.get()
                self._refresh_tree()
                top.destroy()
            entry_widget.bind('<Return>', on_commit)
            entry_widget.bind('<FocusOut>', on_commit)

        # Rotation (Combobox)
        elif col_index == 4:
            top = tk.Toplevel(self.tree)
            top.overrideredirect(True)  # Remove title bar for borderless popup
            popup_height = height + 10
            top.geometry(f"{width}x{popup_height}+{abs_x}+{abs_y}")
            values = ["0", "90", "180", "270"]
            current = str(entry.get("rotation", 0))
            var = tk.StringVar(value=current if current in values else "0")
            opt = tk.OptionMenu(top, var, *values)
            opt.pack(fill="both", expand=True, padx=2, pady=2)
            opt.focus_set()

            def commit_and_close(*args):
                val = var.get()
                try:
                    entry["rotation"] = int(val)
                except Exception:
                    entry["rotation"] = 0
                self._refresh_tree()
                top.destroy()

            # Commit on selection or focus out
            var.trace_add('write', lambda *a: commit_and_close())
            opt.bind('<FocusOut>', lambda e: commit_and_close())
            opt.bind('<Return>', lambda e: commit_and_close())

        # Reverse (Checkbox)
        elif col_index == 5:
            # Add padding to geometry for better centering
            pad_x, pad_y = 4, 2
            popup_width = width + pad_x * 2
            popup_height = height + pad_y * 2
            top = tk.Toplevel(self.tree)
            top.overrideredirect(True)
            top.geometry(f"{popup_width}x{popup_height}+{abs_x-pad_x}+{abs_y-pad_y}")
            var = tk.BooleanVar(value=entry.get("reverse", False))
            frame = ttk.Frame(top)
            frame.pack(fill="both", expand=True)
            cb = ttk.Checkbutton(frame, variable=var, text="", style="TCheckbutton")
            cb.pack(anchor="center", pady=pad_y)
            cb.focus_set()

            def on_commit(event=None):
                entry["reverse"] = var.get()
                self._refresh_tree()
                top.destroy()
            cb.bind('<FocusOut>', on_commit)
            cb.bind('<Return>', on_commit)

        # File name: do nothing
        else:
            return

    def _build_file_buttons(self, parent: ttk.Frame) -> None:
        btn_frame = ttk.Frame(parent)
        btn_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(5, 0))
        for i in range(8):
            btn_frame.columnconfigure(i, weight=1)

        from ttkbootstrap import Button as TBButton
        TBButton(btn_frame, text="Add Files...", command=self.on_add_files, style="WinButton.TButton").grid(row=0, column=0, sticky="w")
        TBButton(btn_frame, text="Add Folder...", command=self.on_add_folder, style="WinButton.TButton").grid(row=0, column=1, sticky="w")
        TBButton(btn_frame, text="Remove Selected", command=self.on_remove_selected, style="WinButton.TButton").grid(row=0, column=2, sticky="w")

        # Move Up button with repeat on hold
        self._move_up_repeat_id = None
        def move_up_press(event=None):
            self.on_move_up()
            self._move_up_repeat_id = btn_move_up.after(200, move_up_press)
        def move_up_release(event=None):
            if self._move_up_repeat_id:
                btn_move_up.after_cancel(self._move_up_repeat_id)
                self._move_up_repeat_id = None
        btn_move_up = TBButton(btn_frame, text="Move Up", style="WinButton.TButton")
        btn_move_up.grid(row=0, column=3, sticky="w")
        btn_move_up.bind('<ButtonPress-1>', move_up_press)
        btn_move_up.bind('<ButtonRelease-1>', move_up_release)

        # Move Down button with repeat on hold
        self._move_down_repeat_id = None
        def move_down_press(event=None):
            self.on_move_down()
            self._move_down_repeat_id = btn_move_down.after(200, move_down_press)
        def move_down_release(event=None):
            if self._move_down_repeat_id:
                btn_move_down.after_cancel(self._move_down_repeat_id)
                self._move_down_repeat_id = None
        btn_move_down = TBButton(btn_frame, text="Move Down", style="WinButton.TButton")
        btn_move_down.grid(row=0, column=4, sticky="w")
        btn_move_down.bind('<ButtonPress-1>', move_down_press)
        btn_move_down.bind('<ButtonRelease-1>', move_down_release)

        TBButton(btn_frame, text="Clear All", command=self.on_clear, style="WinButton.TButton").grid(row=0, column=5, sticky="w")
        TBButton(btn_frame, text="Save List", command=self.on_save_file_list, style="WinButton.TButton").grid(row=0, column=6, sticky="w")
        TBButton(btn_frame, text="Load List", command=self.on_load_file_list, style="WinButton.TButton").grid(row=0, column=7, sticky="w")
    def on_add_folder(self) -> None:
        from core.file_manager import add_files_to_list, SUPPORTED_EXTS
        import os
        folder = filedialog.askdirectory(
            parent=self.root,
            title="Select folder to add files from",
            initialdir=self.settings.last_open_dir or str(ROOT)
        )
        if not folder:
            return
        self.settings.last_open_dir = folder
        self._save_app_settings()
        # Recursively find all supported files
        file_paths = []
        for root, dirs, files in os.walk(folder):
            for fname in files:
                if fname.lower().endswith(SUPPORTED_EXTS):
                    file_paths.append(os.path.join(root, fname))
        if not file_paths:
            messagebox.showinfo("No supported files", "No PDF or image files found in the selected folder.", parent=self.root)
            return
        added, dupes, dupe_names, unsupported, unsupported_names = add_files_to_list(self.files, file_paths)
        msg = []
        if dupes:
            msg.append(f"{dupes} duplicate file(s) skipped: {', '.join(dupe_names)}")
        if unsupported:
            msg.append(f"{unsupported} unsupported file(s) skipped: {', '.join(unsupported_names)}")
        if msg:
            messagebox.showinfo("Some files skipped", "\n".join(msg), parent=self.root)
        self._refresh_tree()
        self._update_status_bar()

    def on_remove_selected(self) -> None:
        from core.file_manager import remove_file
        sels = self.tree.selection()
        if not sels:
            return
        # Remove all selected files, highest index first
        for idx in sorted((int(s) for s in sels), reverse=True):
            remove_file(self.files, idx)
        self._refresh_tree()
        self._update_status_bar()


    def on_move_up(self) -> None:
        from core.file_manager import move_up
        idx = self._get_selected_index()
        if idx is None or idx == 0:
            return
        move_up(self.files, idx)
        self._refresh_tree()
        new_idx = idx - 1
        self.tree.selection_set(str(new_idx))
        self.tree.see(str(new_idx))


    def on_move_down(self) -> None:
        from core.file_manager import move_down
        idx = self._get_selected_index()
        if idx is None or idx == len(self.files) - 1:
            return
        move_down(self.files, idx)
        self._refresh_tree()
        new_idx = idx + 1
        self.tree.selection_set(str(new_idx))
        self.tree.see(str(new_idx))


    def on_clear(self) -> None:
        def show_clear_all_dialog(parent):
            import tkinter as tk
            from ttkbootstrap import Button as TBButton
            dlg = tk.Toplevel(parent)
            dlg.title("Confirm Clear All")
            dlg.transient(parent)
            dlg.grab_set()
            dlg.resizable(False, False)
            # Set dialog size
            width, height = 500, 200
            # Center dialog over parent
            parent.update_idletasks()
            px = parent.winfo_rootx()
            py = parent.winfo_rooty()
            pw = parent.winfo_width()
            ph = parent.winfo_height()
            x = px + (pw // 2) - (width // 2)
            y = py + (ph // 2) - (height // 2)
            dlg.geometry(f"{width}x{height}+{x}+{y}")
            bg = parent.cget('background')
            dlg.configure(bg=bg)
            dlg.columnconfigure(1, weight=1)
            # Icon (reuse warning style)
            try:
                from pathlib import Path
                warning_img = tk.PhotoImage(file=str(Path(__file__).resolve().parent / "images" / "warning.png"))
            except Exception:
                warning_img = None
            if warning_img:
                icon_label = tk.Label(dlg, image=warning_img, bg=bg)
                icon_label.image = warning_img
                icon_label.grid(row=0, column=0, padx=(24, 12), pady=(24, 12), sticky="n")
            else:
                icon_label = tk.Label(dlg, text="!", font=("Segoe UI", 24, "bold"), fg="red", bg=bg)
                icon_label.grid(row=0, column=0, padx=(24, 12), pady=(24, 12), sticky="n")
            # Message
            msg = "Are you sure you want to clear all files from the list?"
            msg_label = tk.Label(
                dlg,
                text=msg,
                font=("Segoe UI", 11),
                bg=bg,
                anchor="w",
                justify="left",
                wraplength=320
            )
            msg_label.grid(row=0, column=1, columnspan=2, sticky="w", padx=(0, 16), pady=(24, 8))
            # Buttons
            result = [None]
            def on_yes():
                result[0] = True
                dlg.destroy()
            def on_no():
                result[0] = False
                dlg.destroy()
            btn_frame = ttk.Frame(dlg)
            btn_frame.grid(row=2, column=0, columnspan=3, pady=(0, 16), sticky="s")
            btn_frame.columnconfigure(0, weight=1)
            btn_frame.columnconfigure(1, weight=1)
            btn_width = 14
            btn_yes = TBButton(btn_frame, text="Yes", command=on_yes, style="WinButton.TButton", width=btn_width)
            btn_yes.grid(row=0, column=0, padx=(0, 12))
            btn_no = TBButton(btn_frame, text="No", command=on_no, style="WinButton.TButton", width=btn_width)
            btn_no.grid(row=0, column=1, padx=(12, 0))
            dlg.protocol("WM_DELETE_WINDOW", on_no)
            dlg.wait_window()
            return result[0]

        if show_clear_all_dialog(self.root):
                from core.file_manager import clear_files
                clear_files(self.files)
                self._refresh_tree()
                self._update_status_bar()

    # -----------------------------------------------------------------------
    # File list callbacks
    # -----------------------------------------------------------------------

    def on_add_files(self) -> None:
        from core.file_manager import add_files_to_list
        initial_dir = self.settings.last_open_dir or str(ROOT)
        paths = filedialog.askopenfilenames(
            parent=self.root,
            title="Select PDF or image files",
            initialdir=initial_dir,
            filetypes=[
                ("PDF and Image files", "*.pdf;*.jpg;*.jpeg;*.png;*.bmp;*.tif;*.tiff"),
                ("PDF files", "*.pdf"),
                ("Image files", "*.jpg;*.jpeg;*.png;*.bmp;*.tif;*.tiff"),
                ("All files", "*.*"),
            ],
        )
        if not paths:
            return

        # Remember last directory
        self.settings.last_open_dir = os.path.dirname(paths[0])
        self._save_app_settings()

        # Use file_manager logic for adding files
        added, dupes, dupe_names, unsupported, unsupported_names = add_files_to_list(self.files, list(paths))

        msg = []
        if dupes:
            msg.append(f"{dupes} duplicate file(s) skipped: {', '.join(dupe_names)}")
        if unsupported:
            msg.append(f"{unsupported} unsupported file(s) skipped: {', '.join(unsupported_names)}")
        # Show a message box if any files could not be loaded (unsupported or unreadable)
        if unsupported or (added == 0 and not dupes):
            # If all files failed, or some unsupported, show the list
            failed_files = unsupported_names.copy()
            if added == 0 and not dupes:
                # All files failed to load (not supported or unreadable)
                failed_files = [os.path.basename(p) for p in paths]
            if failed_files:
                messagebox.showerror(
                    "File(s) could not be loaded",
                    "The following file(s) could not be loaded:\n\n" + "\n".join(failed_files) + "\n\nMake sure they are unencrypted and have\na supported extension (.pdf, .png, etc)",
                    parent=self.root
                )
        elif msg:
            messagebox.showinfo("Some files skipped", "\n".join(msg), parent=self.root)

        self._refresh_tree()
        self._update_status_bar()




    # -----------------------------------------------------------------------
    # Settings notebook
    # -----------------------------------------------------------------------

    def _build_settings_notebook(self, parent: ttk.Frame) -> None:
        print(">>> BUILDING SETTINGS NOTEBOOK")

        nb = ttk.Notebook(parent)
        nb.grid(row=0, column=0, sticky="nsew", pady=(10,10 ))

        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)

        self._build_tab_general(nb)
        self._build_tab_watermark(nb)
        self._build_tab_metadata(nb)
        self._build_tab_scaling(nb)
        self._build_tab_compression(nb)
        self._build_tab_encryption(nb)

    def _build_tab_encryption(self, nb: ttk.Notebook) -> None:
        frame = ttk.Frame(nb, padding=(30, 24, 30, 10))
        nb.add(frame, text="Encryption")

        # Enable encryption checkbox
        self.var_encrypt_enabled = tk.BooleanVar(value=False)
        cb_bg = "#dcdad5"
        cb_fg = "#000000"
        enable_cb = ttk.Checkbutton(frame, text="Encrypt merged file", variable=self.var_encrypt_enabled)
        enable_cb.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        # Variables for passwords
        self.var_encrypt_user_pw = tk.StringVar()
        self.var_encrypt_user_pw2 = tk.StringVar()
        self.var_encrypt_owner_pw = tk.StringVar()
        self.var_encrypt_owner_pw2 = tk.StringVar()

        self._encryption_controls = []



        # Helper to add password entry with eye icon aligned right
        def add_pw_row(label, row, var, padx=(32,0), pady=(0,0)):
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=pady)
            entry_frame = ttk.Frame(frame)
            entry_frame.grid(row=row, column=1, sticky="ew", pady=pady, padx=padx)
            entry_frame.columnconfigure(0, weight=1)
            entry = ttk.Entry(entry_frame, textvariable=var, show="*", width=38)
            entry.grid(row=0, column=0, sticky="ew")
            show = {'state': False}
            def toggle_pw():
                show['state'] = not show['state']
                entry.config(show='' if show['state'] else '*')
                btn.config(text='👁' if not show['state'] else '🙈')
            btn = tk.Button(entry_frame, text='👁', width=2, command=toggle_pw, relief='flat', bd=0, padx=0, pady=0)
            btn.grid(row=0, column=1, sticky="e", padx=(4,0))
            self._encryption_controls.append(entry)
            return entry

        user_pw_entry = add_pw_row("User Password:", 1, self.var_encrypt_user_pw)
        user_pw2_entry = add_pw_row("Retype User Password:", 2, self.var_encrypt_user_pw2)
        # Add extra vertical space before Owner password rows
        owner_pw_entry = add_pw_row("Permissions (Owner) Password:", 4, self.var_encrypt_owner_pw, pady=(24,0))
        owner_pw2_entry = add_pw_row("Retype Owner Password:", 5, self.var_encrypt_owner_pw2)

        # Warning label for mismatched passwords
        self.encrypt_pw_warning = ttk.Label(frame, text="", foreground="red")
        self.encrypt_pw_warning.grid(row=6, column=0, columnspan=2, sticky="w", pady=(5, 0))
        self._encryption_controls.append(self.encrypt_pw_warning)

        frame.columnconfigure(1, weight=1)

        def set_encryption_controls_state(enabled):
            set_widgets_state(self._encryption_controls, enabled)
            if not enabled:
                self.var_encrypt_user_pw.set("")
                self.var_encrypt_user_pw2.set("")
                self.var_encrypt_owner_pw.set("")
                self.var_encrypt_owner_pw2.set("")
                self.encrypt_pw_warning.config(text="")

        def on_encrypt_enabled(*args):
            set_encryption_controls_state(self.var_encrypt_enabled.get())

        self.var_encrypt_enabled.trace_add('write', on_encrypt_enabled)
        set_encryption_controls_state(False)

        def check_passwords(*args):
            if not self.var_encrypt_enabled.get():
                self.encrypt_pw_warning.config(text="")
                return
            user1 = self.var_encrypt_user_pw.get()
            user2 = self.var_encrypt_user_pw2.get()
            owner1 = self.var_encrypt_owner_pw.get()
            owner2 = self.var_encrypt_owner_pw2.get()
            msg = ""
            if user1 or user2:
                if user1 != user2:
                    msg = "User passwords do not match."
            if not msg and (owner1 or owner2):
                if owner1 != owner2:
                    msg = "Owner passwords do not match."
            self.encrypt_pw_warning.config(text=msg)

        self.var_encrypt_user_pw.trace_add('write', check_passwords)
        self.var_encrypt_user_pw2.trace_add('write', check_passwords)
        self.var_encrypt_owner_pw.trace_add('write', check_passwords)
        self.var_encrypt_owner_pw2.trace_add('write', check_passwords)

    # -----------------------------------------------------------------------
    # Settings tabs
    # -----------------------------------------------------------------------

    def _build_tab_general(self, nb: ttk.Notebook) -> None:
        frame = ttk.Frame(nb, padding=(30, 24, 30, 10))
        nb.add(frame, text="General")
        self.var_add_breaker_pages = tk.BooleanVar(value=self.settings.add_breaker_pages)
        self.var_breaker_uniform = tk.BooleanVar(value=self.settings.breaker_uniform_size)
        self.var_delete_blank = tk.BooleanVar(value=self.settings.delete_blank_pages)
        self.var_insert_toc = tk.BooleanVar(value=self.settings.insert_toc)
        self.var_add_filename_bookmarks = tk.BooleanVar(value=self.settings.add_filename_bookmarks)
        # Standardized vertical spacing for all checkboxes
        checkbox_pady = (0, 14)

        # Dark mode toggle
        self.var_dark_mode = tk.BooleanVar(value=getattr(self.settings, 'dark_mode', False))
        def on_dark_mode_toggle():
            self.settings.dark_mode = self.var_dark_mode.get()
            self._save_app_settings()
            # Change theme live using ttkbootstrap Style
            style = tb.Style()
            if self.settings.dark_mode:
                style.theme_use('darkly')
            else:
                style.theme_use('flatly')
        ttk.Checkbutton(frame, text="Dark mode", variable=self.var_dark_mode, command=on_dark_mode_toggle).grid(row=5, column=0, sticky="w", pady=checkbox_pady)

        # Apply dark mode at startup if enabled
        if self.var_dark_mode.get():
            on_dark_mode_toggle()

        cb_bg = "#dcdad5"
        cb_fg = "#000000"
        ttk.Checkbutton(frame, text="Add breaker pages between files", variable=self.var_add_breaker_pages, command=self._on_breaker_pages_toggle).grid(row=0, column=0, sticky="w", pady=checkbox_pady)
        self.breaker_uniform_cb = ttk.Checkbutton(frame, text="Uniform breaker page size", variable=self.var_breaker_uniform)
        self.breaker_uniform_cb.grid(row=1, column=0, sticky="w", padx=24, pady=checkbox_pady)
        self.breaker_uniform_cb.configure(state="normal" if self.var_add_breaker_pages.get() else "disabled")
        ttk.Checkbutton(frame, text="Ignore blank pages in source files when combining", variable=self.var_delete_blank).grid(row=2, column=0, sticky="w", pady=checkbox_pady)
        ttk.Checkbutton(frame, text="Insert a Table of Contents (TOC)", variable=self.var_insert_toc).grid(row=3, column=0, sticky="w", pady=checkbox_pady)
        ttk.Checkbutton(frame, text="Add filename bookmarks", variable=self.var_add_filename_bookmarks).grid(row=4, column=0, sticky="w", pady=checkbox_pady)

    def _on_breaker_pages_toggle(self):
        set_widgets_state([self.breaker_uniform_cb], self.var_add_breaker_pages.get())
        if not self.var_add_breaker_pages.get():
            self.var_breaker_uniform.set(False)

    def _build_tab_watermark(self, nb: ttk.Notebook) -> None:
        # Font color picker callback must be defined before button creation
        self.var_wm_font_color = tk.StringVar(value=getattr(self.settings, 'watermark_font_color', '#000000'))
        def pick_font_color():
            import tkinter.colorchooser
            color = tkinter.colorchooser.askcolor(color=self.var_wm_font_color.get(), parent=self.root)
            if color and color[1]:
                self.var_wm_font_color.set(color[1])
                font_color_btn.config(bg=color[1])
                # Update settings immediately so the correct color is used
                self.settings.watermark_font_color = color[1]
        self.var_wm_safe = tk.BooleanVar(value=self.settings.watermark_safe_mode)
        frame = ttk.Frame(nb, padding=(30, 24, 30, 10))
        nb.add(frame, text="Watermark")

        # --- Controls to be enabled/disabled ---
        self._wm_controls = []

        # (Removed duplicate font size and color widgets; now only single-line widgets are shown below)

        # Safe mode checkbox (spans both columns)

        self.var_wm_enabled = tk.BooleanVar(value=self.settings.watermark_enabled)
        self.var_wm_text = tk.StringVar(value=self.settings.watermark_text)
        self.var_wm_opacity = tk.DoubleVar(value=self.settings.watermark_opacity)
        self.var_wm_rotation = tk.IntVar(value=self.settings.watermark_rotation)
        self.var_wm_position = tk.StringVar(value=self.settings.watermark_position)
        self.var_wm_font_size = tk.IntVar(value=self.settings.watermark_font_size)
        self.var_wm_safe = tk.BooleanVar(value=self.settings.watermark_safe_mode)

        cb_bg = "#dcdad5"
        cb_fg = "#000000"

        for i in range(3):
            frame.grid_columnconfigure(i, weight=0)
        frame.grid_columnconfigure(1, weight=1)

        # --- Controls to be enabled/disabled ---
        self._wm_controls = []

        # Enable watermark checkbox (spans both columns)
        enable_cb = ttk.Checkbutton(frame, text="Enable watermark", variable=self.var_wm_enabled)
        enable_cb.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))

        # Watermark text
        ttk.Label(frame, text="Text:").grid(row=1, column=0, sticky="e", pady=(0, 8), padx=(0, 8))
        wm_text_entry = ttk.Entry(frame, textvariable=self.var_wm_text, width=28)
        wm_text_entry.grid(row=1, column=1, columnspan=2, sticky="we", pady=(0, 8))
        self._wm_controls.append(wm_text_entry)

        # Opacity slider (0-100)
        ttk.Label(frame, text="Opacity:").grid(row=2, column=0, sticky="e", pady=(0, 8), padx=(0, 8))
        self.wm_opacity_value_label = ttk.Label(frame, text=str(int(self.var_wm_opacity.get() * 100)), width=4)
        self.wm_opacity_value_label.grid(row=2, column=2, sticky="w", padx=(8, 0), pady=(0, 8))
        class SnapOpacityScale(ttk.Scale):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                self.bind('<ButtonRelease-1>', self._snap)
                self.bind('<KeyRelease>', self._snap)
            def _snap(self, event=None):
                val = self.get()
                rounded = int(round(val / 5.0) * 5)
                rounded = min(max(rounded, 0), 100)
                if abs(val - rounded) > 0.1:
                    self.set(rounded)
                self.event_generate('<<OpacitySnap>>')

        def update_opacity_label(val):
            rounded = int(round(float(val) / 5.0) * 5)
            rounded = min(max(rounded, 0), 100)
            if self.var_wm_opacity.get() * 100 != rounded:
                self.var_wm_opacity.set(rounded / 100.0)
            self.wm_opacity_value_label.config(text=str(rounded))
            # Only set slider if value differs, and avoid recursion
            if abs(opacity_slider.get() - rounded) > 0.1:
                opacity_slider.set(rounded)

        opacity_slider = SnapOpacityScale(
            frame,
            from_=0,
            to=100,
            orient="horizontal",
            command=update_opacity_label
        )
        opacity_slider.set(int(round(self.var_wm_opacity.get() * 100 / 5.0) * 5))
        opacity_slider.grid(row=2, column=1, sticky="we", pady=(0, 8))
        self._wm_controls.append(opacity_slider)
        self._wm_controls.append(self.wm_opacity_value_label)

        # Add tick marks and descriptive labels under the opacity slider
        opacity_tick_frame = ttk.Frame(frame)
        # Move tick marks label frame very close to the slider
        opacity_tick_frame.grid(row=3, column=1, sticky="ew", pady=(0, 0))
        opacity_tick_frame.columnconfigure(0, weight=1)
        opacity_tick_frame.columnconfigure(1, weight=1)
        opacity_tick_frame.columnconfigure(2, weight=1)
        min_label = ttk.Label(opacity_tick_frame, text="0 - Transparent", anchor="w", font=("Segoe UI", 8))
        min_label.grid(row=0, column=0, sticky="w")
        center_label = ttk.Label(opacity_tick_frame, text="|", anchor="center", font=("Segoe UI", 8))
        center_label.grid(row=0, column=1, sticky="ew")
        max_label = ttk.Label(opacity_tick_frame, text="Opaque - 100", anchor="e", font=("Segoe UI", 8))
        max_label.grid(row=0, column=2, sticky="e")

        # Move all widgets below the slider down by one row
        # Rotation slider (0-359), snap to increments of 5
        ttk.Label(frame, text="Rotation:").grid(row=4, column=0, sticky="e", pady=(0, 8), padx=(0, 8))
        self.wm_rotation_value_label = ttk.Label(frame, text=str(self.var_wm_rotation.get()), width=4)
        self.wm_rotation_value_label.grid(row=4, column=2, sticky="w", padx=(8, 0), pady=(16, 8))
        def update_rotation_label(val):
            rounded = int(round(float(val) / 5.0) * 5)
            rounded = min(max(rounded, 0), 359)
            if self.var_wm_rotation.get() != rounded:
                self.var_wm_rotation.set(rounded)
            self.wm_rotation_value_label.config(text=str(rounded))
            if abs(rotation_slider.get() - rounded) > 0.1:
                rotation_slider.set(rounded)

        class SnapRotationScale(ttk.Scale):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                self.bind('<ButtonRelease-1>', self._snap)
                self.bind('<KeyRelease>', self._snap)
            def _snap(self, event=None):
                val = self.get()
                rounded = int(round(val / 5.0) * 5)
                rounded = min(max(rounded, 0), 359)
                if abs(val - rounded) > 0.1:
                    self.set(rounded)

        rotation_slider = SnapRotationScale(
            frame,
            from_=0,
            to=359,
            orient="horizontal",
            variable=self.var_wm_rotation,
            command=update_rotation_label
        )
        rotation_slider.grid(row=4, column=1, sticky="we", pady=(0, 0))
        # Add descriptive label immediately below rotation slider
        rotation_label_frame = ttk.Frame(frame)
        rotation_label_frame.grid(row=5, column=1, sticky="ew", pady=(0, 0))
        rotation_label_frame.columnconfigure(0, weight=1)
        rotation_desc_label = ttk.Label(
            rotation_label_frame,
            text="0-Horiz   counterclockwise to   Horiz-359",
            font=("Segoe UI", 8),
            foreground="#555555",
            anchor="center"
        )
        rotation_desc_label.grid(row=0, column=0, sticky="ew")
        rotation_slider.set(self.var_wm_rotation.get())
        self._wm_controls.append(rotation_slider)
        self._wm_controls.append(self.wm_rotation_value_label)


        # Position, font size, and font color widgets on the same line
        ttk.Label(frame, text="Position:").grid(row=6, column=0, sticky="e", pady=(24, 8), padx=(0, 8))
        wm_position_combo = ttk.Combobox(
            frame,
            textvariable=self.var_wm_position,
            values=["Top-Left", "Top-Right", "Center", "Bottom-Left", "Bottom-Right"],
            state="readonly",
            width=12
        )
        wm_position_combo.grid(row=6, column=1, sticky="w", pady=(24, 8))
        self._wm_controls.append(wm_position_combo)

        ttk.Label(frame, text="Font size:").grid(row=6, column=2, sticky="e", pady=(24, 8), padx=(16, 8))
        wm_font_entry = ttk.Entry(frame, textvariable=self.var_wm_font_size, width=10)
        wm_font_entry.grid(row=6, column=3, sticky="w", pady=(24, 8))
        ttk.Label(frame, text="Font color:").grid(row=6, column=4, sticky="e", pady=(24, 8), padx=(16, 8))
        font_color_btn = tk.Button(frame, text="Pick Color", command=pick_font_color, bg=self.var_wm_font_color.get(), width=12)
        font_color_btn.grid(row=6, column=5, sticky="w", pady=(24, 8))
        # Ensure button color matches saved color on startup
        font_color_btn.config(bg=self.var_wm_font_color.get())
        self._wm_controls.append(wm_font_entry)
        self._wm_controls.append(font_color_btn)

        # Safe mode checkbox (spans both columns)
        wm_safe_cb = ttk.Checkbutton(frame, text="Dynamically resize to prevent clipping", variable=self.var_wm_safe)
        wm_safe_cb.grid(row=8, column=0, columnspan=3, sticky="w", pady=(0, 8))
        self._wm_controls.append(wm_safe_cb)

        frame.grid_columnconfigure(1, weight=1)

        def set_wm_controls_state(enabled):
            set_widgets_state(self._wm_controls, enabled)
        def on_wm_enabled(*args):
            set_wm_controls_state(self.var_wm_enabled.get())
        self.var_wm_enabled.trace_add('write', lambda *a: on_wm_enabled())
        set_wm_controls_state(self.var_wm_enabled.get())

    def _build_tab_metadata(self, nb: ttk.Notebook) -> None:
        frame = ttk.Frame(nb, padding=(30, 24, 30, 10))
        nb.add(frame, text="Metadata")

        self.var_meta_enabled = tk.BooleanVar(value=self.settings.metadata_enabled)
        self.var_meta_title = tk.StringVar(value=self.settings.pdf_title)
        self.var_meta_author = tk.StringVar(value=self.settings.pdf_author)
        self.var_meta_subject = tk.StringVar(value=self.settings.pdf_subject)
        self.var_meta_keywords = tk.StringVar(value=self.settings.pdf_keywords)

        cb_bg = "#dcdad5"
        cb_fg = "#000000"

        self._meta_controls = []

        ttk.Checkbutton(frame, text="Insert metadata into combined file", variable=self.var_meta_enabled).grid(row=0, column=0, sticky="w", pady=(0, 12))

        ttk.Label(frame, text="Title:").grid(row=1, column=0, sticky="w", pady=(0, 8))
        meta_title_entry = ttk.Entry(frame, textvariable=self.var_meta_title)
        meta_title_entry.grid(row=1, column=1, sticky="ew", pady=(0, 8))
        self._meta_controls.append(meta_title_entry)

        ttk.Label(frame, text="Author:").grid(row=2, column=0, sticky="w", pady=(0, 8))
        meta_author_entry = ttk.Entry(frame, textvariable=self.var_meta_author)
        meta_author_entry.grid(row=2, column=1, sticky="ew", pady=(0, 8))
        self._meta_controls.append(meta_author_entry)

        ttk.Label(frame, text="Subject:").grid(row=3, column=0, sticky="w", pady=(0, 8))
        meta_subject_entry = ttk.Entry(frame, textvariable=self.var_meta_subject)
        meta_subject_entry.grid(row=3, column=1, sticky="ew", pady=(0, 8))
        self._meta_controls.append(meta_subject_entry)

        ttk.Label(frame, text="Keywords:").grid(row=4, column=0, sticky="w", pady=(0, 8))
        meta_keywords_entry = ttk.Entry(frame, textvariable=self.var_meta_keywords)
        meta_keywords_entry.grid(row=4, column=1, sticky="ew", pady=(0, 8))
        self._meta_controls.append(meta_keywords_entry)

        frame.columnconfigure(1, weight=1)

        def set_meta_controls_state(enabled):
            set_widgets_state(self._meta_controls, enabled)
        def on_meta_enabled(*args):
            set_meta_controls_state(self.var_meta_enabled.get())
        self.var_meta_enabled.trace_add('write', lambda *a: on_meta_enabled())
        set_meta_controls_state(self.var_meta_enabled.get())

    def _build_tab_scaling(self, nb: ttk.Notebook) -> None:
        frame = ttk.Frame(nb, padding=(30, 24, 30, 10))
        nb.add(frame, text="Scaling")

        self.var_scale_enabled = tk.BooleanVar(value=self.settings.scaling_enabled)
        self.var_scale_mode = tk.StringVar(value=self.settings.scaling_mode)
        self.var_scale_percent = tk.IntVar(value=self.settings.scaling_percent)

        cb_bg = "#dcdad5"
        cb_fg = "#000000"

        self._scaling_controls = []

        ttk.Checkbutton(frame, text="Enable scaling", variable=self.var_scale_enabled).grid(row=0, column=0, sticky="w", pady=(0, 12))

        ttk.Label(frame, text="Mode:").grid(row=1, column=0, sticky="w", pady=(0, 8))
        scale_mode_combo = ttk.Combobox(
            frame,
            textvariable=self.var_scale_mode,
            values=["Fit", "Fill", "Percent"],
            state="readonly",
        )
        scale_mode_combo.grid(row=1, column=1, sticky="w", pady=(0, 8))
        self._scaling_controls.append(scale_mode_combo)

        ttk.Label(frame, text="Percent:").grid(row=2, column=0, sticky="w", pady=(0, 8))
        scale_percent_entry = ttk.Entry(frame, textvariable=self.var_scale_percent, width=10)
        scale_percent_entry.grid(row=2, column=1, sticky="w", pady=(0, 8))
        self._scaling_controls.append(scale_percent_entry)

        frame.columnconfigure(1, weight=1)

        def set_scaling_controls_state(enabled):
            set_widgets_state(self._scaling_controls, enabled)
        def on_scaling_enabled(*args):
            set_scaling_controls_state(self.var_scale_enabled.get())
        self.var_scale_enabled.trace_add('write', lambda *a: on_scaling_enabled())
        set_scaling_controls_state(self.var_scale_enabled.get())

    def _build_tab_compression(self, nb: ttk.Notebook) -> None:
        frame = ttk.Frame(nb, padding=(30, 24, 30, 10))
        nb.add(frame, text="Compression")

        # Add compression info note
        note = ttk.Label(frame, text="Note: Higher compression levels result in lower quality images.", foreground="red", font=("Segoe UI", 9, "italic"))
        note.grid(row=0, column=2, sticky="w", padx=(16,0), pady=(0, 12))

        self.var_comp_enabled = tk.BooleanVar(value=self.settings.compression_enabled)
        self.var_comp_level = tk.StringVar(value=getattr(self.settings, 'compression_level', 'Medium'))

        def on_comp_enabled_toggle(*args):
            enabled = self.var_comp_enabled.get()
            set_widgets_state([comp_level_combo], enabled)

        cb_bg = "#dcdad5"
        cb_fg = "#000000"
        ttk.Checkbutton(frame, text="Compress merged PDF", variable=self.var_comp_enabled, command=on_comp_enabled_toggle).grid(row=0, column=0, sticky="w", pady=(0, 12))

        ttk.Label(frame, text="Compression Level:").grid(row=1, column=0, sticky="w", pady=(0, 8))
        comp_level_combo = ttk.Combobox(frame, textvariable=self.var_comp_level, values=["Low", "Medium", "High", "Maximum"], state="readonly", width=12)
        comp_level_combo.grid(row=1, column=1, sticky="w", padx=(0, 8), pady=(0, 8))

        frame.columnconfigure(1, weight=1)

        # Set initial state
        on_comp_enabled_toggle()
        self.var_comp_enabled.trace_add('write', lambda *args: on_comp_enabled_toggle())



    # -----------------------------------------------------------------------
    # File list helpers
    # -----------------------------------------------------------------------

    def _refresh_tree(self) -> None:
        import datetime
        from core.file_manager import SUPPORTED_EXTS
        def human_size(num, suffix="B"):
            for unit in ["", "K", "M", "G", "T", "P", "E", "Z"]:
                if abs(num) < 1024.0:
                    return f"{num:3.1f}{unit}{suffix}"
                num /= 1024.0
            return f"{num:.1f}Y{suffix}"

        # Define image extensions
        image_exts = ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.tif')

        # Setup tag for blue text if not already present
        if not self.tree.tag_has('imagefile'):
            self.tree.tag_configure('imagefile', foreground='blue')

        self.tree.delete(*self.tree.get_children())

        # Remove any existing instructional label
        if hasattr(self, '_filelist_instruction_label') and self._filelist_instruction_label:
            self._filelist_instruction_label.destroy()
            self._filelist_instruction_label = None

        if not self.files:
            # Hide Treeview and show instructional label
            self.tree.grid_remove()
            parent = getattr(self, '_filelist_frame', self.tree.master)
            try:
                bg = parent.cget('background')
            except Exception:
                bg = "white"
            self._filelist_instruction_label = tk.Label(
                parent,
                text="*** Quickstart ***\n\nClick 'Add Files' , 'Add Folder' or 'Load File List' to specify individual files to be combined.\n\nSet combining options using Options & Settings.\n\nSpecify the output filename for the merged file.\n\nFinally, click 'Merge' to merge the files.",
                bg=bg,
                anchor="center",
                justify="center"
            )
            self._filelist_instruction_label.grid(row=0, column=0, sticky="nsew")
            return
        else:
            # Show Treeview and remove instructional label if present
            self.tree.grid()
            if hasattr(self, '_filelist_instruction_label') and self._filelist_instruction_label:
                self._filelist_instruction_label.destroy()
                self._filelist_instruction_label = None

        for idx, entry in enumerate(self.files):
            path = entry["path"]
            try:
                stat = os.stat(path)
                size = human_size(stat.st_size)
                mtime = datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            except Exception:
                size = "-"
                mtime = "-"
            ext = os.path.splitext(path)[1].lower()
            tags = ()
            is_image = ext in image_exts
            if is_image:
                tags = ('imagefile',)
            # Show 'N/A' for Pages and Reverse for images
            page_range_val = entry.get("page_range", "All") if not is_image else "N/A"
            reverse_val = ("\u2713" if entry.get("reverse", False) else "\u00D7") if not is_image else "-"
            self.tree.insert(
                "",
                "end",
                iid=str(idx),
                values=(
                    os.path.basename(path),
                    size,
                    mtime,
                    page_range_val,
                    entry.get("rotation", 0),
                    reverse_val,
                ),
                tags=tags
            )

    # -----------------------------------------------------------------------
    # Collect all settings into MergeOptions
    # -----------------------------------------------------------------------

    def _collect_options(self) -> MergeOptions:
        # Sync UI → settings
        self.settings.add_breaker_pages = self.var_add_breaker_pages.get()
        self.settings.breaker_uniform_size = self.var_breaker_uniform.get()
        self.settings.delete_blank_pages = self.var_delete_blank.get()
        self.settings.insert_toc = self.var_insert_toc.get()
        self.settings.add_filename_bookmarks = self.var_add_filename_bookmarks.get()

        self.settings.compression_enabled = self.var_comp_enabled.get()
        self.settings.compression_level = self.var_comp_level.get()

        self.settings.watermark_enabled = self.var_wm_enabled.get()
        self.settings.watermark_text = self.var_wm_text.get()
        self.settings.watermark_opacity = float(self.var_wm_opacity.get())
        self.settings.watermark_rotation = int(self.var_wm_rotation.get())
        self.settings.watermark_position = self.var_wm_position.get()
        self.settings.watermark_font_size = int(self.var_wm_font_size.get())
        self.settings.watermark_safe_mode = self.var_wm_safe.get()

        self.settings.metadata_enabled = self.var_meta_enabled.get()
        self.settings.pdf_title = self.var_meta_title.get()
        self.settings.pdf_author = self.var_meta_author.get()
        self.settings.pdf_subject = self.var_meta_subject.get()
        self.settings.pdf_keywords = self.var_meta_keywords.get()

        self.settings.scaling_enabled = self.var_scale_enabled.get()
        self.settings.scaling_mode = self.var_scale_mode.get()
        self.settings.scaling_percent = int(self.var_scale_percent.get())

        # Encryption options from UI
        encrypt_enabled = getattr(self, 'var_encrypt_enabled', None)
        encrypt_user_pw = getattr(self, 'var_encrypt_user_pw', None)
        encrypt_owner_pw = getattr(self, 'var_encrypt_owner_pw', None)
        encryption = {
            'encrypt_enabled': bool(encrypt_enabled.get()) if encrypt_enabled else False,
            'encrypt_user_pw': encrypt_user_pw.get() if encrypt_user_pw else "",
            'encrypt_owner_pw': encrypt_owner_pw.get() if encrypt_owner_pw else "",
        }

        self._save_app_settings()

        return MergeOptions(
            add_breaker_pages=self.settings.add_breaker_pages,
            breaker_uniform_size=self.settings.breaker_uniform_size,

            delete_blank_pages=self.settings.delete_blank_pages,
            insert_toc=self.settings.insert_toc,

            compression_enabled=self.settings.compression_enabled,
            compression_level=self.settings.compression_level,

            watermark_enabled=self.settings.watermark_enabled,
            watermark_text=self.settings.watermark_text,
            watermark_opacity=self.settings.watermark_opacity,
            watermark_rotation=self.settings.watermark_rotation,
            watermark_position=self.settings.watermark_position,
            watermark_font_size=self.settings.watermark_font_size,
            watermark_safe_mode=self.settings.watermark_safe_mode,
            watermark_font_color=self.settings.watermark_font_color,

            metadata_enabled=self.settings.metadata_enabled,
            pdf_title=self.settings.pdf_title,
            pdf_author=self.settings.pdf_author,
            pdf_subject=self.settings.pdf_subject,
            pdf_keywords=self.settings.pdf_keywords,

            scaling_enabled=self.settings.scaling_enabled,
            scaling_mode=self.settings.scaling_mode,
            scaling_percent=self.settings.scaling_percent,

            # Encryption
            encrypt_enabled=encryption['encrypt_enabled'],
            encrypt_user_pw=encryption['encrypt_user_pw'],
            encrypt_owner_pw=encryption['encrypt_owner_pw'],

            # Bookmarks
            add_filename_bookmarks=self.settings.add_filename_bookmarks,
        )

    # -----------------------------------------------------------------------
    # Merge button
    # -----------------------------------------------------------------------

    def on_merge_clicked(self) -> None:
        if not self.files:
            messagebox.showwarning("No files", "Please add at least one file.", parent=self.root)
            return

        output_path = self.output_var.get().strip()
        # Ensure .pdf extension
        if output_path and not output_path.lower().endswith('.pdf'):
            output_path += '.pdf'
            self.output_var.set(output_path)
            self.settings.output_filename = output_path
            self._save_app_settings()
        if not output_path:
            messagebox.showwarning("No output", "Please choose an output file.", parent=self.root)
            return

        # Check if output file exists before starting merge
        if os.path.exists(output_path):
            warning_img_path = str(Path(__file__).resolve().parent / "images" / "warning.png")
            result = show_custom_dialog(
                self.root,
                title="File Exists",
                message=f"The file '{output_path}' already exists.\nDo you want to overwrite it?",
                icon=warning_img_path,
                buttons=["Overwrite", "Cancel"],
                default="Cancel",
                cancel="Cancel",
                width=520,
                height=210
            )
            if result != "Overwrite":
                return

        entries: List[FileEntry] = [
            FileEntry(
                path=e["path"],
                rotation=e.get("rotation", 0),
                page_range=e.get("page_range", "All"),
                reverse=e.get("reverse", False),
            )
            for e in self.files
        ]

        options = self._collect_options()

        # Prevent encryption if passwords do not match
        if options.encrypt_enabled:
            user_pw1 = self.var_encrypt_user_pw.get()
            user_pw2 = self.var_encrypt_user_pw2.get()
            owner_pw1 = self.var_encrypt_owner_pw.get()
            owner_pw2 = self.var_encrypt_owner_pw2.get()
            if user_pw1 != user_pw2:
                messagebox.showerror("Password Mismatch", "User passwords do not match.", parent=self.root)
                return
            if owner_pw1 != owner_pw2:
                messagebox.showerror("Password Mismatch", "Owner passwords do not match.", parent=self.root)
                return

        # Progress dialog + background thread
        self._progress_dialog = ProgressDialog(self.root, title="Merging PDFs...")
        if self._progress_dialog:
            self._progress_dialog.set_filename("Preparing to merge ...")

        def progress_callback(idx, total, filename):
            if self._progress_dialog:
                self._progress_dialog.set_filename(os.path.basename(filename) if filename else "")

        def cancel_callback():
            return self._progress_dialog and self._progress_dialog.cancelled

        self._merge_thread = threading.Thread(
            target=self._run_merge,
            args=(entries, output_path, options, progress_callback, cancel_callback),
            daemon=True,
        )
        self._merge_thread.start()
        self._poll_merge_thread()




    # -----------------------------------------------------------------------
    # Background merge worker
    # -----------------------------------------------------------------------

    def _run_merge(self, entries, output_path, options, progress_callback=None, cancel_callback=None):
        try:
            if progress_callback is not None or cancel_callback is not None:
                merge_files(entries, output_path, options, progress_callback=progress_callback, cancel_callback=cancel_callback)
            else:
                merge_files(entries, output_path, options)
            self._merge_error = None
        except Exception as e:
            self._merge_error = e

    # -----------------------------------------------------------------------
    # Polling loop for merge thread
    # -----------------------------------------------------------------------

    def _poll_merge_thread(self):
        if self._merge_thread is None:
            return

        if self._merge_thread.is_alive():
            self.root.after(100, self._poll_merge_thread)
            return

        # Thread finished
        if self._progress_dialog:
            self._progress_dialog.close()
            self._progress_dialog = None

        if self._merge_error:
            err_msg = str(self._merge_error)
            file_name = None
            import re
            # Try to extract filename for PDF or image errors
            pdf_match = re.match(r"([\w\-.\\/:]+\.pdf)", err_msg)
            img_match = re.search(r"image file '([\w\-.\\/:]+)'", err_msg)
            if pdf_match:
                file_name = pdf_match.group(1)
            elif img_match:
                file_name = img_match.group(1)
            # Also try to extract image file from error message
            if file_name:
                messagebox.showerror("Merge failed", f"File: {file_name}\n\n{err_msg}", parent=self.root)
            else:
                messagebox.showerror("Merge failed", err_msg, parent=self.root)
        else:
            self._show_merge_done_dialog()

        self._merge_thread = None

    def _show_merge_done_dialog(self):
        output_path = self.output_var.get().strip()
        file_count = len(self.files)
        try:
            file_size = os.path.getsize(output_path)
        except Exception:
            file_size = None

        def human_size(size):
            if size is None:
                return "Unknown size"
            for unit in ['bytes', 'KB', 'MB', 'GB', 'TB']:
                if size < 1024.0:
                    return f"{size:.1f} {unit}" if unit != 'bytes' else f"{size:,} bytes"
                size /= 1024.0
            return f"{size:.1f} PB"

        size_str = human_size(file_size)
        filename = os.path.basename(output_path)
        fullpath = os.path.abspath(output_path)

        dlg = tk.Toplevel(self.root)
        dlg.title("Done")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)
        # Match ProgressDialog color scheme
        bg = self.root.cget('background')
        dlg.configure(bg=bg)

        # Set icon if available
        icon_path = Path(__file__).resolve().parent / "pdfcombinericon.ico"
        if icon_path.exists():
            try:
                dlg.iconbitmap(str(icon_path))
            except Exception:
                pass

        # Set a taller, more compact size
        width, height = 440, 300
        dlg.geometry(f"{width}x{height}")

        # Center the dialog in the parent window
        self.root.update_idletasks()
        root_x = self.root.winfo_rootx()
        root_y = self.root.winfo_rooty()
        root_w = self.root.winfo_width()
        root_h = self.root.winfo_height()
        x = root_x + (root_w // 2) - (width // 2)
        y = root_y + (root_h // 2) - (height // 2)
        dlg.geometry(f"{width}x{height}+{x}+{y}")

        bg = dlg.cget("background")
        dlg.columnconfigure(1, weight=1)
        dlg.rowconfigure(0, weight=1)
        dlg.rowconfigure(1, weight=0)
        dlg.rowconfigure(2, weight=0)

        # Add check.png icon
        check_img_path = Path(__file__).resolve().parent / "images" / "check.png"
        check_label = None
        if check_img_path.exists():
            try:
                check_img = tk.PhotoImage(file=str(check_img_path))
                check_label = tk.Label(dlg, image=check_img, bg=bg)
                check_label.image = check_img
            except Exception:
                check_label = tk.Label(dlg, text="✔", font=("Segoe UI", 24, "bold"), fg="green", bg=bg)
        else:
            check_label = tk.Label(dlg, text="✔", font=("Segoe UI", 24, "bold"), fg="green", bg=bg)
        check_label.grid(row=0, column=0, rowspan=2, padx=(24, 12), pady=(24, 12), sticky="n")

        info = (
            f"Merged PDF created successfully.\n\n"
            f"Files combined: {file_count}\n"
            f"Output size: {size_str}\n"
            f"Saved as: {fullpath}"
        )
        msg_label = tk.Label(
            dlg,
            text=info,
            font=("Segoe UI", 10),
            bg=bg,
            anchor="w",
            justify="left",
            wraplength=300
        )
        msg_label.grid(row=0, column=1, columnspan=2, sticky="w", padx=(0, 16), pady=(24, 24))

        def open_pdf():
            try:
                if sys.platform.startswith("win"):
                    os.startfile(output_path)
                elif sys.platform == "darwin":
                    import subprocess
                    subprocess.Popen(["open", output_path])
                else:
                    import subprocess
                    subprocess.Popen(["xdg-open", output_path])
            except Exception as e:
                messagebox.showerror("Open Failed", f"Could not open PDF:\n{e}", parent=dlg)

        from ttkbootstrap import Button as TBButton
        btn_frame = tk.Frame(dlg, bg=bg, bd=0, highlightthickness=0)
        btn_frame.grid(row=2, column=0, columnspan=3, pady=(0, 18), sticky="sew")
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)
        open_btn = TBButton(btn_frame, text="Open PDF", command=open_pdf, width=14, style="WinButton.TButton")
        open_btn.grid(row=0, column=0, padx=(10, 8), sticky="e")
        close_btn = TBButton(btn_frame, text="Close", command=dlg.destroy, width=14, style="WinButton.TButton")
        close_btn.grid(row=0, column=1, padx=(8, 10), sticky="w")
        dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)

    def on_save_file_list(self) -> None:
        import json
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title="Save file list as...",
            defaultextension=".pdflist",
            filetypes=[("PDF List files", "*.pdflist")],
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.files, f, indent=2)
        except Exception as e:
            messagebox.showerror("Save Failed", f"Could not save file list:\n{e}", parent=self.root)

    def on_load_file_list(self) -> None:
        import json
        from core.file_manager import add_files_to_list
        path = filedialog.askopenfilename(
            parent=self.root,
            title="Load file list",
            filetypes=[("PDF List files", "*.pdflist")],
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            # Validate loaded data: only keep dicts with a path and supported extension
            from core.file_manager import SUPPORTED_EXTS
            valid_files = []
            unsupported_files = []
            for entry in loaded:
                if isinstance(entry, dict) and "path" in entry:
                    if any(entry["path"].lower().endswith(e) for e in SUPPORTED_EXTS):
                        valid_files.append({
                            "path": entry["path"],
                            "page_range": entry.get("page_range", "All"),
                            "rotation": entry.get("rotation", 0),
                            "reverse": entry.get("reverse", False),
                        })
                    else:
                        unsupported_files.append(entry["path"])
            self.files.clear()
            self.files.extend(valid_files)
            if unsupported_files:
                messagebox.showinfo("Some files skipped", "Unsupported file(s) skipped:\n" + "\n".join(unsupported_files), parent=self.root)
            self._refresh_tree()
            self._update_status_bar()
        except Exception as e:
            messagebox.showerror("Load Failed", f"Could not load file list:\n{e}", parent=self.root)

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    print("CWD:", os.getcwd())
    print("Loaded settings.py from:", CORE_DIR / "settings.py")

    # Enable DPI awareness on Windows
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per-monitor DPI aware
    except Exception:
        pass

    # Load settings to determine theme
    settings = None
    try:
        settings = load_settings(CONFIG_PATH)
    except Exception:
        pass
    theme = "flatly"
    if settings and getattr(settings, 'dark_mode', False):
        theme = "darkly"

    import time
    from PIL import Image, ImageTk

    # Splash screen setup
    splash = tk.Tk()
    splash.overrideredirect(True)
    splash_img_path = os.path.join(ROOT, "images", "splashscreen.png")
    splash_img = None
    splash_label = None
    try:
        img = Image.open(splash_img_path)
        splash_img = ImageTk.PhotoImage(img)
        splash_label = tk.Label(splash, image=splash_img)
        splash_label.pack()
        splash.update_idletasks()
        w = img.width
        h = img.height
        sw = splash.winfo_screenwidth()
        sh = splash.winfo_screenheight()
        x = (sw // 2) - (w // 2)
        y = (sh // 2) - (h // 2)
        splash.geometry(f"{w}x{h}+{x}+{y}")
    except Exception:
        splash.destroy()
        splash = None

    if splash:
        splash.after(2000, splash.destroy)
        splash.update()
        splash.deiconify()
        splash.mainloop()

    root = tb.Window(themename=theme)
    root.withdraw()  # Hide window during setup
    icon_path = Path(__file__).resolve().parent / "pdfcombinericon.ico"
    if icon_path.exists():
        try:
            root.iconbitmap(str(icon_path))
        except Exception:
            pass

    # Center the main window
    root.update_idletasks()
    width = 800
    height = 600
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")

    CombinePDFsUI(root)
    root.deiconify()  # Show window after centering
    root.mainloop()


if __name__ == "__main__":
    main()