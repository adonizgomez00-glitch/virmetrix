# Build VirMetrix for Windows with PyInstaller.
# Run in PowerShell from the project root.

$ErrorActionPreference = "Stop"

if (-not (Test-Path ".venv")) {
    python -m venv .venv
}
.venv\Scripts\Activate.ps1
pip install -r requirements-dev.txt
pip install pyinstaller

pyinstaller --clean --noconfirm virmetrix.spec

Write-Host "Build complete: dist\VirMetrix\VirMetrix.exe"
Write-Host "To create an installer, use Inno Setup or NSIS over dist\VirMetrix."