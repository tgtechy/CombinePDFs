# PowerShell script to compute and display SHA256 hash of the CombinePDFs installer
# This script is called after the installer build completes

param(
    [string]$InstallerPath = "dist\CombinePDFs.exe"
)

# Check if installer exists
if (-not (Test-Path $InstallerPath)) {
    Write-Host "Error: Installer not found at $InstallerPath" -ForegroundColor Red
    exit 1
}

# Compute SHA256 hash
try {
    $hash = (Get-FileHash $InstallerPath -Algorithm SHA256).Hash
    $fileSize = (Get-Item $InstallerPath).Length
    
    # Format file size
    if ($fileSize -lt 1MB) {
        $fileSizeStr = "{0:F2} KB" -f ($fileSize / 1KB)
    } else {
        $fileSizeStr = "{0:F2} MB" -f ($fileSize / 1MB)
    }
    
    # Display results
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Installer SHA256 Verification Hash" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "File:     $(Split-Path $InstallerPath -Leaf)" -ForegroundColor Yellow
    Write-Host "Size:     $fileSizeStr" -ForegroundColor Yellow
    Write-Host "Hash:     $hash" -ForegroundColor Green
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Save hash to file
    $outputDir = Split-Path $InstallerPath
    $hashFile = Join-Path $outputDir "$(Split-Path $InstallerPath -Leaf).sha256"
    
    # Save in formatted format with header
    $hashContent = @"
File:     $(Split-Path $InstallerPath -Leaf)
Size:     $fileSizeStr
Hash:     $hash
"@
    $hashContent | Out-File $hashFile -Encoding UTF8
    
    Write-Host "Hash saved to: $hashFile" -ForegroundColor Gray
    Write-Host ""
    
    exit 0
}
catch {
    Write-Host "Error computing hash: $_" -ForegroundColor Red
    exit 1
}
