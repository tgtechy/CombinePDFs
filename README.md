CombinePDFs — Build instructions

This project is a small Tkinter app to combine PDF files.

Build a Windows executable (one-file) using PyInstaller

Prerequisites
- Python 3.8+ installed and on PATH
- Recommended: create and activate a virtual environment

Quick build (from project root):

1) Install dependencies (in venv):

```powershell
python -m pip install -r requirements.txt
```

2) Build using the helper script (Windows):

```powershell
.\build_exe.bat
```

What the script does
- Installs/updates PyInstaller
- Includes `pdfcombinericon.png` or `pdf_icon.ico` if present
- Runs PyInstaller to create `dist\CombinePDFs.exe` (single-file, windowed)

Manual PyInstaller command

If you prefer to run PyInstaller manually, use:

```powershell
pyinstaller --noconfirm --onefile --windowed --name CombinePDFs combine_pdfs.py
```

If you have icon files to include (from project root):

```powershell
pyinstaller --noconfirm --onefile --windowed --name CombinePDFs --add-data "pdfcombinericon.png;." combine_pdfs.py
```

Notes
- The generated exe will be in the `dist` folder.
- Large dependencies (PyMuPDF, Pillow) can increase exe size.
- Test the exe on a clean Windows machine for missing DLLs.
