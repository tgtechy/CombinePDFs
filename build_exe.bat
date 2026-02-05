@echo off
REM Build CombinePDFs into a single Windows executable using PyInstaller
REM Run this from the project root (double-click or run from PowerShell/CMD)

python -m pip install --upgrade pyinstaller

REM Build using the spec file which includes icon and data configurations.
pyinstaller --noconfirm CombinePDFs.spec

REM Post-build: use rcedit to embed the icon into the EXE
if exist dist\CombinePDFs.exe (
    if exist combine_pdfs.ico (
        rcedit --set-icon combine_pdfs.ico dist\CombinePDFs.exe
    )
)





echo Build finished. Output executable is in the "dist" folder as dist\CombinePDFs.exe
pause