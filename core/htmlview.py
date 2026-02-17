__all__ = ["HtmlHelpViewer"]
import tkinter as tk
from tkinter import ttk
import os

try:
    from tkinterweb import HtmlFrame
except ImportError:
    raise ImportError("Please install tkinterweb: pip install tkinterweb")


class HtmlHelpViewer(ttk.Frame):
    """
    Drop-in HTML help viewer that loads instructions.html automatically.

    Usage:
        viewer = HtmlHelpViewer(parent, initial_theme="light",
                                html_path="instructions.html")
        viewer.pack(fill="both", expand=True)
        viewer.set_theme("dark")  # optional
    """

    def __init__(self, parent, initial_theme="light",
                 html_path="instructions.html", *args, **kwargs):
        super().__init__(parent, *args, **kwargs)

        self.current_theme = initial_theme.lower()
        self.html_path = html_path
        self._base_html = ""

        # Toolbar removed (theme is now set programmatically)

        # HTML viewer
        self.browser = HtmlFrame(self, horizontal_scrollbar="auto", messages_enabled=False)
        self.browser.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    # ---------- Public API ----------

    def set_theme(self, theme: str):
        """Switch between light/dark themes."""
        theme = theme.lower()
        if theme not in ("light", "dark"):
            return
        self.current_theme = theme
        if self._base_html:
            themed_html = self._inject_theme_css(self._base_html, theme)
            try:
                self.browser.load_html(themed_html)
            except AttributeError:
                self.browser.set_content(themed_html)

    # ---------- Internal helpers ----------

    def load_html(self, html: str):
        """Public method to load raw HTML into the viewer."""
        self._base_html = html
        themed_html = self._inject_theme_css(html, self.current_theme)
        try:
            self.browser.load_html(themed_html)
        except AttributeError:
            self.browser.set_content(themed_html)

        # _on_theme_change removed (theme is now set programmatically)

    def _load_html_file(self):
        if not os.path.exists(self.html_path):
            self._base_html = f"<h1>Missing File</h1><p>{self.html_path} not found.</p>"
        else:
            with open(self.html_path, "r", encoding="utf-8") as f:
                raw_html = f.read()

            # NEW: rewrite image paths
            raw_html = self._rewrite_image_paths(raw_html)

            self._base_html = raw_html

        themed_html = self._inject_theme_css(self._base_html, self.current_theme)

        try:
            self.browser.load_html(themed_html)
        except AttributeError:
            self.browser.set_content(themed_html)

    def _inject_theme_css(self, html: str, theme: str) -> str:
        """Injects CSS for light/dark mode into the HTML."""
        light_css = """
        body {
            background-color: #ffffff;
            color: #000000;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
            margin: 12px;
        }
        a { color: #0066cc; }
        code, pre {
            font-family: Consolas, "Courier New", monospace;
            background-color: #f4f4f4;
            padding: 2px 4px;
            border-radius: 3px;
        }
        h1, h2, h3, h4 { color: #222222; }
        """

        dark_css = """
        body {
            background-color: #1e1e1e;
            color: #e0e0e0;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
            margin: 12px;
        }
        a { color: #6aa9ff; }
        code, pre {
            font-family: Consolas, "Courier New", monospace;
            background-color: #2d2d2d;
            padding: 2px 4px;
            border-radius: 3px;
        }
        h1, h2, h3, h4 { color: #ffffff; }
        hr {
            border: none;
            border-top: 1px solid #444444;
        }
        """

        css = dark_css if theme == "dark" else light_css
        style_block = f"<style>{css}</style>"

        lower_html = html.lower()

        if "<head>" in lower_html:
            idx = lower_html.index("<head>") + len("<head>")
            return html[:idx] + style_block + html[idx:]

        if "<html>" in lower_html:
            idx = lower_html.index("<html>") + len("<html>")
            return html[:idx] + "<head>" + style_block + "</head>" + html[idx:]

        # No HTML structure → wrap it
        return f"<html><head>{style_block}</head><body>{html}</body></html>"
    
    def _rewrite_image_paths(self, html: str) -> str:
        """Rewrites relative image paths to absolute paths (PyInstaller-safe)."""
        import sys

        # Determine base path (normal run or PyInstaller bundle)
        if hasattr(sys, "_MEIPASS"):
            base = sys._MEIPASS
        else:
            base = os.path.abspath(".")

        images_path = os.path.join(base, "images").replace("\\", "/")

        # Rewrite all occurrences of images/
        return html.replace("images/", images_path + "/")