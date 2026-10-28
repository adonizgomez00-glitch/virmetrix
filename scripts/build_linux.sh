#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")/.."

if [ ! -d ".venv" ]; then
    python3 -m venv .venv
fi
source .venv/bin/activate
pip install -r requirements-dev.txt
pip install pyinstaller

pyinstaller --clean --noconfirm virmetrix.spec
echo "Build complete: dist/VirMetrix/VirMetrix"