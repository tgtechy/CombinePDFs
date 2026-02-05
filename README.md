# CombinePDFs

A program to combine two or more PDF files into one PDF.

## Build Instructions

This project is a Tkinter app to combine PDF files with drag-and-drop reordering.

### Build a Windows executable using PyInstaller

**Prerequisites**
- Python 3.8+ installed and on PATH
- Recommended: create and activate a virtual environment

**Quick build (from project root):**

1) Install dependencies (in venv):

```powershell
python -m pip install -r requirements.txt
```

2) Build using the helper script (Windows):

```powershell
.\build_exe.bat
```

**What the script does**
- Installs/updates PyInstaller
- Includes `pdfcombinericon.png` and `pdfcombinericon.ico`
- Runs PyInstaller to create `dist\CombinePDFs.exe` (single-file, windowed)

**Manual PyInstaller command**

If you prefer to run PyInstaller manually, use:

```powershell
pyinstaller --noconfirm CombinePDFs.spec
```

**Notes**
- The generated exe will be in the `dist` folder.
- Large dependencies (PyMuPDF, Pillow) can increase exe size.
- Test the exe on a clean Windows machine for missing DLLs.
