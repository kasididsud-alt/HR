$ErrorActionPreference = "Stop"

Set-Location -Path $PSScriptRoot

Write-Host "Building SalaryCalc portable package..." -ForegroundColor Cyan

$py = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $py)) {
  throw "Missing .venv\Scripts\python.exe. Please create the virtual environment first."
}

& $py tools\build_portable.py

Write-Host "Done. Output: dist\SalaryCalc_Portable" -ForegroundColor Green
