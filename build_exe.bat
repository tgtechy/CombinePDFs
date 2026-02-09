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

echo.
echo Embedding icon into executable...

REM Embed icon using rcedit
if exist dist\CombinePDFs.exe (
    if exist pdfcombinericon.ico (
        rcedit --set-icon pdfcombinericon.ico dist\CombinePDFs.exe
        echo Icon embedded.
    )
)
echo.

echo Building installer with Inno Setup...

REM Try to find and run Inno Setup compiler
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" CombinePDFs.iss
) else if exist "C:\Program Files\Inno Setup 6\ISCC.exe" (
    "C:\Program Files\Inno Setup 6\ISCC.exe" CombinePDFs.iss
) else (
    echo WARNING: Inno Setup not found at expected locations
    echo Please run Inno Setup manually to compile CombinePDFs.iss
)
echo.

echo Computing SHA256 checksum of final installer...

REM Run the hash computation script on the final installer
REM The .iss output path is determined by OutputDir and OutputBaseFilename
if exist installer\CombinePDFsInstaller_1.5.0.exe (
    powershell -NoProfile -ExecutionPolicy Bypass -File compute_hash.ps1 -InstallerPath "installer\CombinePDFsInstaller_1.5.0.exe"
) else (
    echo WARNING: Final installer not found. Hash computation skipped.
)
echo.

echo Build finished.
pause