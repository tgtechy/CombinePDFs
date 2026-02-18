# PDFCombiner: Merge Several PDF Documents into a Single PDF

A free, standalone Windows 10/11 application for combining two or more PDF files into a single document. It includes powerful options such as page extraction, rotation, bookmarks, encryption, and more. Nothing is ever uploaded or downloaded — all processing happens locally to protect your privacy. After getting tired of searching for a clean, free, ad‑free, non‑web‑based tool, I decided to build one myself.

<p align="left">
  <img src="images/splashscreen_small.png" width="400">
</p>

<p align="left">
  <img src="images/inputtab.png" width="400">
</p>
<p align="left">
  <img src="images/optionstab1.png" width="400">
</p>

# Windows Installation and Use
Click the 'Releases' link on this page then click the PDFCombinerInstaller_X.Y.Z.exe file to download it.
You may get warnings since it is not digitally signed. You can compare the SHA hash of the exe to the one shown in the release notes to ensure the exe has not been tampered with.

To bypass the Windows warning if it appears when trying to install/run, click 'More info', then 'Run anyway' as shown below (pics are from Win11, Win10 may be a little different):

<img src="images/windows-protected-your-pc-click1.png?v=3" width="300">

<img src="images/windows-protected-your-pc-click2.png?v=3" width="300">

If you want to make a donation to support further development, [Donate via PayPal](https://www.paypal.me/tgtechdevshop)

If you would like to build the exe from scratch rather than download the .exe file in the releases section ...

## Build Instructions (if you don't want to use the provided pre-compiled .exe)

This project is a Python/Tkinter app to combine PDF files.

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
