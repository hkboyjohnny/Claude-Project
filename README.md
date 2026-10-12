# Sinochem Holdings: sector peer comparison

Peer comparison of Sinochem Holdings' Energy, Agriculture and Rubber businesses against listed peers (FY2025, with H1 2026 read-through).

| File | What it is |
|---|---|
| `Sinochem_Peer_Comparison_Report.md` | Written report: performance commentary and peer-comparison conclusions per sector |
| `Sinochem_Peer_Comparison.xlsx` | Workbook: **Position** (editable one-glance ranking tables), scorecard, sector matrices (live formulas), H1 2026 momentum, group context, standardised inputs, FX, sources |
| `charts/` | One positioning chart per sector (EBITDA margin, EBITDA/assets, leverage) |
| `scripts/` | Fetch, standardise and rebuild pipeline (`scripts/run_all.sh`) |
| `data/` | `dataset.json` (standardised inputs), `layout.json`, and `raw/` (statements as fetched) |
| `report_src/` | Report template; tables are filled from the workbook by `scripts/build_report.py` |

Rebuild: `scripts/run_all.sh`, then `python3 scripts/build_report.py`. Needs Python 3 with `openpyxl` and `matplotlib`, and LibreOffice for formula recalculation.

Sinochem Holdings is unlisted and publishes no sector P&L, so each sector is anchored on group entities with filed statements (Sinochem Energy; Syngenta AG, ADAMA, Yangnong, Sinofert; Pirelli, Aeolus Tyre). See the report's method section for definitions and limits.
