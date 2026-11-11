# Genera el instalador:  dist\Kardexpress-Instalador-<versión>.exe
# Uso (desde esta carpeta):  powershell -ExecutionPolicy Bypass -File build.ps1
#
# Pasos: pruebas -> ejecutable (PyInstaller) -> instalador (Inno Setup 6).
# Requisito único fuera de Python: Inno Setup 6  (winget install JRSoftware.InnoSetup)
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not (Test-Path .venv)) {
    python -m venv .venv
    .\.venv\Scripts\python.exe -m pip install --upgrade pip
}
.\.venv\Scripts\pip.exe install -q -r requirements.txt pyinstaller pytest

# 1) Pruebas
.\.venv\Scripts\python.exe -m pytest -q
if ($LASTEXITCODE -ne 0) { throw "Las pruebas fallaron; no se genera el instalador." }

# 2) Ejecutable
.\.venv\Scripts\pyinstaller.exe --noconfirm --clean --windowed `
    --name Kardexpress `
    --icon assets\kardexpress.ico `
    --add-data "assets\logo.png;assets" `
    --add-data "assets\kardexpress.ico;assets" `
    --collect-data customtkinter `
    --collect-all tkinterdnd2 `
    --exclude-module numpy `
    --exclude-module pytest `
    run.py
if ($LASTEXITCODE -ne 0) { throw "PyInstaller falló." }

# 3) Instalador
$version = (.\.venv\Scripts\python.exe -c "import kardexpress; print(kardexpress.__version__)").Trim()
$iscc = @(
    "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe",
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) { throw "No se encontró Inno Setup 6. Instálalo con:  winget install JRSoftware.InnoSetup" }

& $iscc /Q "/DAppVersion=$version" installer\kardexpress.iss
if ($LASTEXITCODE -ne 0) { throw "Inno Setup falló." }

Write-Host "`nListo: dist\Kardexpress-Instalador-$version.exe" -ForegroundColor Green
