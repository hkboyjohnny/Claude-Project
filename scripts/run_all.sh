#!/usr/bin/env bash
# Rebuild everything from raw sources. Requires: python3 (openpyxl, matplotlib), LibreOffice (soffice) for formula recalculation.
set -euo pipefail
cd "$(dirname "$0")/.."
python3 -I scripts/fetch_eastmoney.py --all      # A-share / HK / US statements -> data/raw
python3 -I scripts/fetch_sec.py                  # SEC XBRL facts -> data/raw
python3 -I scripts/build_dataset.py              # standardise -> data/dataset.json
python3 -I scripts/build_workbook.py             # formulas workbook (no cached values)
TMP="$(mktemp -d)"
soffice --headless --calc --convert-to xlsx --outdir "$TMP" Sinochem_Peer_Comparison.xlsx >/dev/null 2>&1
cp "$TMP/Sinochem_Peer_Comparison.xlsx" Sinochem_Peer_Comparison.xlsx   # same workbook, formulas + cached values
python3 -I scripts/make_charts.py                # charts/*.png from workbook values
python3 -I scripts/report_tables.py              # markdown tables used in the report
