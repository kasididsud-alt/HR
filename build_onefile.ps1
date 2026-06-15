$ErrorActionPreference = "Stop"

Set-Location -Path $PSScriptRoot

Write-Host "Preparing SalaryCalc Windows build..." -ForegroundColor Cyan

$py = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"

function Try-CreateVenv {
  param(
    [string]$Command,
    [string[]]$CommandArgs
  )

  if (-not (Get-Command $Command -ErrorAction SilentlyContinue)) {
    return $false
  }

  & $Command @CommandArgs
  return (($LASTEXITCODE -eq 0) -and (Test-Path $py))
}

function Run-Native {
  param(
    [string]$Command,
    [string[]]$CommandArgs
  )

  & $Command @CommandArgs
  if ($LASTEXITCODE -ne 0) {
    throw "$Command failed with exit code $LASTEXITCODE"
  }
}

if (-not (Test-Path $py)) {
  $created = Try-CreateVenv "py" @("-3.11", "-m", "venv", ".venv")
  if (-not $created) {
    $created = Try-CreateVenv "py" @("-3", "-m", "venv", ".venv")
  }
  if (-not $created) {
    $created = Try-CreateVenv "python" @("-m", "venv", ".venv")
  }
}

if (-not (Test-Path $py)) {
  throw "Missing .venv\Scripts\python.exe. Python 3.10+ is required."
}

Write-Host "Installing build dependencies..." -ForegroundColor Cyan
Run-Native $py @("-m", "pip", "install", "--upgrade", "pip")
Run-Native $py @("-m", "pip", "install", "-r", "requirements.txt", "pyinstaller")

Write-Host "Building SalaryCalc portable package..." -ForegroundColor Cyan
Run-Native $py @("tools\build_portable.py")

Write-Host "Done. Output: dist\SalaryCalc_Portable.zip" -ForegroundColor Green
