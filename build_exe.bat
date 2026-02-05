@echo off
REM Build CombinePDFs into a single Windows executable using PyInstaller
REM Run this from the project root (double-click or run from PowerShell/CMD)

REM Activate virtual environment if it exists
if exist venv\Scripts\activate.bat (
    call venv\Scripts\activate.bat
)

python -m pip install --upgrade pyinstaller

REM Build using the spec file which includes icon and data configurations.
pyinstaller --noconfirm CombinePDFs.spec

REM Post-build: use rcedit to embed the icon into the EXE
if exist dist\CombinePDFs.exe (
    if exist pdfcombinericon.ico (
        rcedit --set-icon pdfcombinericon.ico dist\CombinePDFs.exe
    )
)





echo Build finished. Output executable is in the "dist" folder as dist\CombinePDFs.exe
pause