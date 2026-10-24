# Bouwt AccessExporter.exe (één bestand, Windows) met PyInstaller.
# Builds AccessExporter.exe (single file, Windows) with PyInstaller.
# Gebruik / usage:  .\build.ps1            (gebruikt .venv als die bestaat / uses .venv when present)
param([string]$Python = "")

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot
if (-not $Python) {
    $Python = if (Test-Path ".venv\Scripts\python.exe") { ".venv\Scripts\python.exe" } else { "python" }
}

& $Python -m pip install --quiet pyinstaller
if ($LASTEXITCODE -ne 0) { throw "pip install pyinstaller failed" }

& $Python -m PyInstaller --noconfirm --clean --onefile --windowed `
    --name AccessExporter `
    --icon "access_exporter\assets\access-exporter.ico" `
    --add-data "access_exporter\assets;access_exporter\assets" `
    --hidden-import xlsxwriter --hidden-import openpyxl `
    access_exporter_pyqt6.py
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed" }

$exe = Join-Path $PSScriptRoot "dist\AccessExporter.exe"
$hash = (Get-FileHash $exe -Algorithm SHA256).Hash
Write-Host "Built $exe"
Write-Host "SHA256 $hash"
