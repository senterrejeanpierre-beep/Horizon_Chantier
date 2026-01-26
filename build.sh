#!/bin/bash
set -e
cd "$(dirname "$0")"

source .venv/bin/activate

rm -rf build dist

pyinstaller \
  --clean \
  --noconfirm \
  --onedir \
  --console \
  --name Devis \
  --add-data "DATA:DATA" \
  --add-data "excel:excel" \
  --add-data "soumission_chapelette.xlsx:." \
  app_devis.py
