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
echo Computing SHA256 checksum (before icon embedding)...

REM Create temporary PowerShell script
echo $hash = ^(Get-FileHash 'dist\CombinePDFs.exe' -Algorithm SHA256^).Hash > temp_hash.ps1
echo Write-Host "SHA256: $hash" >> temp_hash.ps1
echo $hash ^| Out-File 'dist\SHA256.txt' >> temp_hash.ps1

REM Run the script
powershell -NoProfile -ExecutionPolicy Bypass -File temp_hash.ps1

REM Clean up
del temp_hash.ps1
echo Checksum saved to dist\SHA256.txt
echo.

REM Embed icon using rcedit
if exist dist\CombinePDFs.exe (
    if exist pdfcombinericon.ico (
        echo Embedding icon into executable...
        rcedit --set-icon pdfcombinericon.ico dist\CombinePDFs.exe
        echo Icon embedded.
    )
)
echo.





echo Build finished. Output executable is in the "dist" folder as dist\CombinePDFs.exe
pause