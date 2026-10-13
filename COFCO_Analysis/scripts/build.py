"""Build the COFCO analysis pack: data CSVs, charts (PNG) and the Excel workbook.

Run from the COFCO_Analysis folder:  python3 scripts/build.py
Every number lives in DATA below with its source tag; the report quotes these tables.
Source tags map to entries in resources/SOURCES.md.
"""
import csv
import os

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.colors import LinearSegmentedColormap  # noqa: E402
from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(ROOT, "data")
CHART_DIR = os.path.join(ROOT, "charts")

# ---------------------------------------------------------------------------
# Palette (reference categorical palette, light mode, fixed slot order)
# ---------------------------------------------------------------------------
S1, S2, S3, S4 = "#2a78d6", "#eb6834", "#1baf7a", "#eda100"
MUTED = "#a8a69f"
INK, INK2, GRID, SURF = "#0b0b0b", "#52514e", "#e4e3df", "#fcfcfb"
NEG, POS = "#e34948", "#2a78d6"

plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 10,
    "axes.edgecolor": GRID, "axes.labelcolor": INK2, "axes.titlecolor": INK,
    "axes.titlesize": 11.5, "axes.titleweight": "bold", "axes.titlelocation": "left",
    "xtick.color": INK2, "ytick.color": INK2, "axes.grid": True, "grid.color": GRID,
    "grid.linewidth": 0.8, "axes.axisbelow": True, "figure.facecolor": SURF,
    "axes.facecolor": SURF, "axes.spines.top": False, "axes.spines.right": False,
    "legend.frameon": False, "legend.labelcolor": INK2,
})

# ---------------------------------------------------------------------------
# DATA. Each table: columns, rows, source tags, notes.
# Units are stated in column names. "n/a" = not disclosed / not found.
# ---------------------------------------------------------------------------
DATA = {
    "group_financials": {
        "title": "COFCO Corporation (consolidated) - headline financials, RMB bn",
        "cols": ["year", "total_revenue", "total_profit_pre_tax", "net_profit", "net_profit_to_parent",
                 "total_assets", "liabilities_to_assets_pct", "total_debt", "ebitda", "debt_to_ebitda_x"],
        "rows": [
            [2021, 664.95, 23.81, None, None, 686.01, 70.39, 224.54, 39.99, 5.62],
            [2022, 741.44, 22.85, None, None, 695.57, 69.01, 214.36, 40.26, 5.32],
            [2023, 692.10, 21.17, None, None, 730.65, 65.35, 231.64, 39.48, 5.87],
            [2024, 619.05, 11.78, 5.01, 1.87, 693.76, 63.31, None, None, None],
            [2025, 589.06, 12.05, 5.89, 2.41, 721.60, 64.55, None, None, None],
        ],
        "src": ["G1", "G2", "G3"],
        "note": ("2021-2023 from Lianhe ratings (as originally reported). 2024-2025 from the 2025 bond annual report; "
                 "2024 is restated for an accounting policy change (originally RMB 635.0bn revenue / 8.2bn pre-tax profit). "
                 "2024 debt/EBITDA as originally reported: debt 200.6bn, EBITDA 25.9bn, 7.73x."),
    },
    "group_segments": {
        "title": "COFCO Corporation - operating revenue and gross margin by segment, RMB bn",
        "cols": ["segment", "rev_2023", "gm_2023_pct", "rev_2024", "gm_2024_pct", "rev_2025", "gm_2025_pct"],
        "rows": [
            ["Agri trading & processing (grain, oilseed, sugar, cotton, meat)", 600.66, 5.43, 543.20, 5.87, 518.25, 5.82],
            ["Food & packaging (China Foods/Coca-Cola, tea, liquor)", 30.48, 42.79, 28.36, 35.39, 27.95, 33.71],
            ["Hotels & property (Joy City)", 36.90, 26.42, 35.91, 21.41, 31.02, 31.17],
            ["Finance & other", 4.63, 14.89, 2.39, 36.37, 2.38, 35.48],
        ],
        "src": ["G1", "G3"],
        "note": "Operating revenue excludes insurance/interest income. Mengniu (dairy) is equity-accounted, so it is not in these lines.",
    },
    "group_segment_gp_2025": {
        "title": "COFCO Corporation - 2025 revenue share vs gross-profit share",
        "cols": ["segment", "revenue_bn", "gross_profit_bn", "revenue_share_pct", "gp_share_pct"],
        "rows": [],  # derived below
        "src": ["G3"],
        "note": "Gross profit = revenue - cost of sales by segment (bond annual report table).",
    },
    "cofco_intl": {
        "title": "COFCO International - revenue (US$ bn) and volumes (million tonnes)",
        "cols": ["year", "revenue_usd_bn", "volume_mt"],
        "rows": [[2020, 33.0, 131.0], [2021, 48.0, 133.0], [2022, 53.3, 127.0],
                 [2023, 50.1, 121.7], [2024, 38.5, 108.4], [2025, 34.5, 100.4]],
        "src": ["C1", "C2", "C3", "C4"],
        "note": "2020-2021 volume is total traded volume; 2022+ is third-party sales volume. Profit is not disclosed.",
    },
    "traders": {
        "title": "Global grain traders - revenue (US$ bn) and net income (US$ bn)",
        "cols": ["company", "rev_2021", "rev_2022", "rev_2023", "rev_2024", "rev_2025",
                 "ni_2021", "ni_2022", "ni_2023", "ni_2024", "ni_2025"],
        "rows": [
            ["ADM", 85.25, 101.56, 93.94, 85.53, 80.27, 2.71, 4.34, 3.48, 1.80, 1.08],
            ["Bunge", 59.15, 67.23, 59.54, 53.11, 70.33, 2.08, 1.61, 2.24, 1.14, 0.82],
            ["Louis Dreyfus (LDC)", 49.6, 59.9, 50.6, 50.6, 53.2, 0.70, 1.01, 1.01, 0.73, 0.65],
            ["Wilmar", 65.79, 73.40, 67.16, 67.38, 70.42, 1.89, 2.40, 1.52, 1.17, 1.41],
            ["COFCO International", 48.0, 53.3, 50.1, 38.5, 34.5, None, None, None, None, None],
        ],
        "src": ["T1", "T2", "T3", "T4", "C1"],
        "note": ("2025 and 2024 figures checked against 2025 releases; 2021-2023 compiled from company annual reports. "
                 "Bunge 2025 includes Viterra from July 2025. Wilmar 2025 net profit includes a US$1.14bn one-off gain; core net profit US$1.28bn."),
    },
    "soy_imports": {
        "title": "China soybean imports, million tonnes (customs)",
        "cols": ["year", "imports_mt"],
        "rows": [[2019, 88.51], [2020, 100.33], [2021, 96.52], [2022, 91.08], [2023, 99.41], [2024, 105.03], [2025, 111.83]],
        "src": ["M1"],
        "note": "2025 = 111.83mt (+6.5%), a record. Earlier years from GACC annual releases.",
    },
    "joycome": {
        "title": "COFCO Joycome - key operating and financial data",
        "cols": ["period", "revenue_rmb_bn", "profit_before_ba_rmb_bn", "profit_after_ba_rmb_bn",
                 "hog_asp_rmb_per_kg", "hog_volume_m_heads", "fresh_pork_kt"],
        "rows": [
            ["2019", 11.08, 0.37, 1.52, 16.16, 1.985, 146],
            ["2020", 18.92, 2.91, 4.02, 32.33, 2.046, 109],
            ["2021", 13.23, 2.35, -0.48, 18.26, 3.437, 166],
            ["2022", 12.90, -0.32, 0.36, 17.86, 4.103, 231],
            ["2023", 17.24, 0.15, -0.03, 14.55, 5.200, 245],
            ["2024", 16.33, 0.28, 0.55, 16.53, 3.560, 262],
            ["2025", 18.58, -0.29, -0.74, 13.39, None, 341],
            ["H1 2025", 8.96, 0.20, 0.32, 14.59, None, 150],
            ["H1 2026", 8.74, -1.10, -1.56, 10.27, None, 197],
        ],
        "src": ["J1", "J2", "J3", "J4", "J5", "J6"],
        "note": "Profit for the period (incl. minorities). BA = biological asset fair value adjustment. Hog volume no longer disclosed after 2024.",
    },
    "joycome_segments": {
        "title": "COFCO Joycome - segment results, RMB m",
        "cols": ["segment", "FY2024", "FY2025", "H1_2025", "H1_2026"],
        "rows": [
            ["Hog production", 258.6, -159.1, 243.9, -993.7],
            ["Feed", 202.4, 157.9, 55.8, 54.8],
            ["Fresh pork", -57.2, -5.6, -4.3, 24.4],
            ["Processed meat", 35.5, 6.1, 6.2, -12.2],
            ["Meat import", 58.2, 19.9, 15.8, 24.0],
        ],
        "src": ["J1", "J2"],
        "note": "Segment results before fair value adjustments, corporate items and finance costs.",
    },
    "pork_peers": {
        "title": "China hog producers - comparison",
        "cols": ["company", "np_2024_rmb_bn", "np_2025_rmb_bn", "np_H1_2026_rmb_bn",
                 "hogs_sold_2025_m", "asp_H1_2025", "asp_H1_2026"],
        "rows": [
            ["Muyuan", 17.88, 15.49, -6.2, 77.98, 14.4, 10.4],
            ["Wens", 9.28, 5.27, -4.37, 40.48, 14.93, 10.50],
            ["New Hope", 0.47, -1.78, None, 17.55, None, None],
            ["COFCO Joycome", 0.54, -0.75, -1.55, None, 14.59, 10.27],
        ],
        "src": ["P1", "P2", "P3", "J1", "J2"],
        "note": ("Net profit attributable to shareholders (after BA for Joycome). Muyuan H1 2026 = midpoint of its "
                 "RMB 5.7-6.7bn loss guidance; Muyuan H1 2025 ASP derived from the ~28% decline it reported."),
    },
    "sow_herd": {
        "title": "China breeding sow inventory, million head (NBS / MARA)",
        "cols": ["date", "sows_m", "normal_level_m"],
        "rows": [["2022 peak", 43.90, 41.0], ["End-2024", 40.79, 39.0], ["End-2025", 39.61, 39.0], ["Jun-2026", 37.80, 37.5]],
        "src": ["J1", "J2"],
        "note": "Normal-level reserve as quoted by MARA guidance at each date (41.0m pre-2024 guideline). End-2024 derived from the -2.9% change.",
    },
    "mengniu": {
        "title": "China Mengniu Dairy - revenue and profit, RMB bn",
        "cols": ["period", "revenue", "np_attrib", "np_adjusted", "gross_margin_pct"],
        "rows": [
            ["2019", 79.03, 4.11, None, None], ["2020", 76.04, 3.53, None, None],
            ["2021", 88.14, 5.03, None, None], ["2022", 92.59, 5.30, None, None],
            ["2023", 98.62, 4.81, None, None], ["2024", 88.68, 0.105, 4.435, 39.6],
            ["2025", 82.24, 1.545, 3.960, 39.9], ["H1 2025", 41.57, 2.046, None, 41.7],
            ["H1 2026", 44.79, 2.371, None, 40.8],
        ],
        "src": ["D1", "D2", "D3", "D4"],
        "note": "Adjusted profit excludes the Bellamy's impairment (2024: RMB 3.98bn) and associate one-offs (Mengniu's own definition).",
    },
    "mengniu_segments": {
        "title": "Mengniu revenue by segment, RMB m",
        "cols": ["segment", "FY2024", "FY2025", "H1_2025", "H1_2026"],
        "rows": [
            ["Liquid milk", 73066, 64939, 32192, 33865],
            ["Ice cream", 5175, 5393, 3879, 4207],
            ["Cheese", 4320, 5266, 2374, 3148],
            ["Milk formula & others", 6114, 6647, 3122, 3575],
        ],
        "src": ["D1", "D2"],
        "note": "H1 'Milk formula & others' = total less the three named segments.",
    },
    "dairy_peers": {
        "title": "Mengniu vs Yili - revenue and attributable net profit, RMB bn",
        "cols": ["year", "mengniu_rev", "yili_rev", "mengniu_np", "yili_np"],
        "rows": [
            [2019, 79.03, 90.22, 4.11, 6.93], [2020, 76.04, 96.89, 3.53, 7.08],
            [2021, 88.14, 110.60, 5.03, 8.70], [2022, 92.59, 123.17, 5.30, 9.43],
            [2023, 98.62, 126.18, 4.81, 10.43], [2024, 88.68, 115.78, 0.105, 8.45],
            [2025, 82.24, 115.93, 1.545, 11.57],
        ],
        "src": ["D1", "D5", "D6"],
        "note": "H1 2026: Yili revenue RMB 64.5bn (+4.1%), net profit RMB 5.76bn (-20.0%, goodwill impairment on Ausnutria); Mengniu +7.8% / +15.9%.",
    },
}

# derived: 2025 segment gross profit
_seg25 = [("Agri trading & processing", 518.245, 488.074), ("Food & packaging", 27.950, 18.527),
          ("Hotels & property", 31.022, 21.354), ("Finance & other", 2.383, 1.538)]
_tot_r = sum(r for _, r, _ in _seg25)
_tot_g = sum(r - c for _, r, c in _seg25)
DATA["group_segment_gp_2025"]["rows"] = [
    [n, round(r, 1), round(r - c, 1), round(100 * r / _tot_r, 1), round(100 * (r - c) / _tot_g, 1)] for n, r, c in _seg25
]

# Analyst cycle scoring (judgement, -2 deep trough .. +2 strong upcycle)
CYCLE = {
    "years": ["2019", "2020", "2021", "2022", "2023", "2024", "2025", "H1 2026", "2027E"],
    "rows": {
        "Grains & oilseeds": [0, 1, 2, 2, 1, -1, -1, -1, 0],
        "Pork (Joycome)": [1, 2, 0, -1, -1, 1, -1, -2, 1],
        "Dairy (Mengniu)": [1, 0, 1, 1, 0, -2, -1, 0, 1],
    },
}


def write_csvs():
    os.makedirs(DATA_DIR, exist_ok=True)
    for key, t in DATA.items():
        with open(os.path.join(DATA_DIR, f"{key}.csv"), "w", newline="") as f:
            w = csv.writer(f)
            w.writerow(t["cols"])
            for r in t["rows"]:
                w.writerow(["" if v is None else v for v in r])
    with open(os.path.join(DATA_DIR, "cycle_scores.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["sector"] + CYCLE["years"])
        for k, v in CYCLE["rows"].items():
            w.writerow([k] + v)


# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------
def _save(fig, name, source):
    fig.text(0.01, 0.005, f"Source: {source}", fontsize=7.5, color=INK2, ha="left", va="bottom")
    fig.tight_layout(rect=(0, 0.03, 1, 1))
    fig.savefig(os.path.join(CHART_DIR, name), dpi=160)
    plt.close(fig)


def _bar_labels(ax, bars, fmt="{:.0f}", inside=False):
    for b in bars:
        h = b.get_height()
        if h is None:
            continue
        va = "bottom" if h >= 0 else "top"
        ax.annotate(fmt.format(h), (b.get_x() + b.get_width() / 2, h), xytext=(0, 3 if h >= 0 else -3),
                    textcoords="offset points", ha="center", va=va, fontsize=8.5, color=INK2)


def charts():
    os.makedirs(CHART_DIR, exist_ok=True)
    bar_kw = dict(edgecolor=SURF, linewidth=2)

    # 1 Group revenue and pre-tax profit (small multiples, no dual axis)
    g = DATA["group_financials"]["rows"]
    yrs = [str(r[0]) for r in g]
    fig, (a, b) = plt.subplots(1, 2, figsize=(10, 3.8))
    _bar_labels(a, a.bar(yrs, [r[1] for r in g], color=S1, **bar_kw))
    a.set_title("Total revenue, RMB bn")
    _bar_labels(b, b.bar(yrs, [r[2] for r in g], color=S1, **bar_kw), "{:.1f}")
    b.set_title("Total pre-tax profit, RMB bn")
    for ax in (a, b):
        ax.grid(axis="x", visible=False)
    _save(fig, "01_group_revenue_profit.png", "Lianhe Ratings 2024/2025; COFCO 2025 bond annual report (2024 restated)")

    # 2 Segment revenue share vs gross profit share 2025
    s = DATA["group_segment_gp_2025"]["rows"]
    names = [r[0] for r in s]
    fig, ax = plt.subplots(figsize=(10, 3.6))
    y = range(len(names))
    h = 0.38
    b1 = ax.barh([i + h / 2 for i in y], [r[3] for r in s], h, color=S1, label="Share of revenue", **bar_kw)
    b2 = ax.barh([i - h / 2 for i in y], [r[4] for r in s], h, color=S2, label="Share of gross profit", **bar_kw)
    for bars in (b1, b2):
        for bb in bars:
            ax.annotate(f"{bb.get_width():.0f}%", (bb.get_width(), bb.get_y() + bb.get_height() / 2),
                        xytext=(4, 0), textcoords="offset points", va="center", fontsize=8.5, color=INK2)
    ax.set_yticks(list(y), names)
    ax.invert_yaxis()
    ax.set_xlim(0, 100)
    ax.grid(axis="y", visible=False)
    ax.set_title("2025: agri is 89% of revenue but only 60% of gross profit")
    ax.legend(loc="lower right")
    _save(fig, "02_segment_revenue_vs_gp_2025.png", "COFCO 2025 bond annual report, segment table")

    # 3 Segment gross margins 2023-2025
    seg = DATA["group_segments"]["rows"]
    short = ["Agri", "Food", "Property", "Finance"]
    fig, ax = plt.subplots(figsize=(10, 3.6))
    w = 0.26
    for j, (lbl, col, idx) in enumerate([("2023", MUTED, 2), ("2024", S3, 4), ("2025", S1, 6)]):
        bars = ax.bar([i + (j - 1) * w for i in range(4)], [r[idx] for r in seg], w, color=col, label=lbl, **bar_kw)
        _bar_labels(ax, bars, "{:.1f}")
    ax.set_xticks(range(4), short)
    ax.set_title("Gross margin by segment, % - agri is thin-margin, high-volume")
    ax.grid(axis="x", visible=False)
    ax.legend(ncol=3, loc="upper left")
    _save(fig, "03_segment_gross_margin.png", "Lianhe Ratings 2025; COFCO 2025 bond annual report")

    # 4 COFCO International revenue & volume
    c = DATA["cofco_intl"]["rows"]
    yrs = [str(r[0]) for r in c]
    fig, (a, b) = plt.subplots(1, 2, figsize=(10, 3.6))
    _bar_labels(a, a.bar(yrs, [r[1] for r in c], color=S1, **bar_kw), "{:.1f}")
    a.set_title("COFCO International revenue, US$ bn")
    _bar_labels(b, b.bar(yrs, [r[2] for r in c], color=S3, **bar_kw))
    b.set_title("Volume handled, million tonnes")
    b.set_ylim(0, 150)
    for ax in (a, b):
        ax.grid(axis="x", visible=False)
    _save(fig, "04_cofco_international.png", "COFCO International sustainability reports via Reuters / World Grain / company site")

    # 5 Trader peers: revenue 2025 and net margin trend
    t = DATA["traders"]["rows"]
    fig, (a, b) = plt.subplots(1, 2, figsize=(11, 4))
    order = sorted(t, key=lambda r: r[5])
    cols = [S1 if r[0] == "COFCO International" else MUTED for r in order]
    bars = a.barh([r[0] for r in order], [r[5] for r in order], color=cols, **bar_kw)
    for bb in bars:
        a.annotate(f"{bb.get_width():.1f}", (bb.get_width(), bb.get_y() + bb.get_height() / 2), xytext=(4, 0),
                   textcoords="offset points", va="center", fontsize=8.5, color=INK2)
    a.set_title("2025 revenue, US$ bn")
    a.grid(axis="y", visible=False)
    years = ["2021", "2022", "2023", "2024", "2025"]
    for r, col in zip(t[:4], [S1, S2, S3, S4]):
        m = [100 * r[6 + i] / r[1 + i] for i in range(5)]
        b.plot(years, m, color=col, lw=2, marker="o", ms=5, label=r[0])
    b.set_title("Net margin, % - the industry peaked in 2022")
    b.set_xlim(-0.2, 4.2)
    b.legend(loc="upper right", fontsize=8)
    _save(fig, "05_trader_peers.png", "ADM 8-K, Bunge Q4-25 release, LDC, Wilmar results; COFCO Intl (profit not disclosed)")

    # 6 China soybean imports
    s = DATA["soy_imports"]["rows"]
    fig, ax = plt.subplots(figsize=(10, 3.4))
    bars = ax.bar([str(r[0]) for r in s], [r[1] for r in s], color=[MUTED] * 6 + [S1], **bar_kw)
    _bar_labels(ax, bars, "{:.1f}")
    ax.set_ylim(0, 125)
    ax.grid(axis="x", visible=False)
    ax.set_title("China soybean imports, million tonnes - record 112mt in 2025")
    _save(fig, "06_china_soybean_imports.png", "General Administration of Customs (via SunSirs / Reuters)")

    # 7 Joycome hog price vs profit
    j = DATA["joycome"]["rows"]
    per = [r[0] for r in j if not r[0].startswith("H1")] + ["H1 2026"]
    jj = [r for r in j if r[0] in per]
    fig, (a, b) = plt.subplots(1, 2, figsize=(11, 3.9))
    a.plot(per, [r[4] for r in jj], color=S2, lw=2, marker="o", ms=6)
    for x, r in zip(per, jj):
        a.annotate(f"{r[4]:.1f}", (x, r[4]), xytext=(0, 7), textcoords="offset points", ha="center", fontsize=8, color=INK2)
    a.set_title("Joycome hog selling price, RMB/kg")
    a.set_ylim(0, 38)
    vals = [r[2] for r in jj]
    bars = b.bar(per, vals, color=[POS if v >= 0 else NEG for v in vals], **bar_kw)
    _bar_labels(b, bars, "{:.2f}")
    b.axhline(0, color=INK2, lw=0.8)
    b.set_title("Profit before bio-asset adjustment, RMB bn")
    b.set_ylim(-1.6, 3.4)
    for ax in (a, b):
        ax.tick_params(axis="x", labelrotation=30)
        ax.grid(axis="x", visible=False)
    _save(fig, "07_joycome_price_vs_profit.png", "COFCO Joycome annual / interim results announcements (HKEX)")

    # 8 Joycome segment results
    js = DATA["joycome_segments"]["rows"]
    fig, ax = plt.subplots(figsize=(10, 3.8))
    w = 0.2
    periods = ["FY2024", "FY2025", "H1_2025", "H1_2026"]
    for k, (p, col) in enumerate(zip(periods, [MUTED, S3, S4, S1])):
        ax.bar([i + (k - 1.5) * w for i in range(len(js))], [r[1 + k] for r in js], w, color=col,
               label=p.replace("_", " "), **bar_kw)
    ax.axhline(0, color=INK2, lw=0.8)
    ax.set_xticks(range(len(js)), [r[0] for r in js])
    ax.set_title("Joycome segment results, RMB m - hog farming drives almost all of the swing")
    ax.grid(axis="x", visible=False)
    ax.legend(ncol=4, loc="lower right")
    _save(fig, "08_joycome_segments.png", "COFCO Joycome FY2025 and H1 2026 results, segment note")

    # 9 Pork peers: ASP H1 25 vs H1 26 and profit
    pp = DATA["pork_peers"]["rows"]
    fig, (a, b) = plt.subplots(1, 2, figsize=(11, 3.8))
    peers = [r for r in pp if r[5] is not None]
    x = range(len(peers))
    a.bar([i - 0.2 for i in x], [r[5] for r in peers], 0.4, color=MUTED, label="H1 2025", **bar_kw)
    bars = a.bar([i + 0.2 for i in x], [r[6] for r in peers], 0.4, color=S2, label="H1 2026", **bar_kw)
    _bar_labels(a, bars, "{:.1f}")
    a.set_xticks(list(x), [r[0] for r in peers])
    a.set_title("Hog selling price, RMB/kg: ~-29% for everyone")
    a.legend(loc="upper right")
    a.set_ylim(0, 18)
    names = [r[0] for r in pp]
    w = 0.26
    for k, (idx, lbl, col) in enumerate([(1, "2024", MUTED), (2, "2025", S3), (3, "H1 2026", S2)]):
        vals = [r[idx] if r[idx] is not None else 0 for r in pp]
        b.bar([i + (k - 1) * w for i in range(len(pp))], vals, w, color=col, label=lbl, **bar_kw)
    b.axhline(0, color=INK2, lw=0.8)
    b.set_xticks(range(len(pp)), names, fontsize=8.5)
    b.set_title("Net profit, RMB bn")
    b.legend(ncol=3, fontsize=8, loc="upper right")
    for ax in (a, b):
        ax.grid(axis="x", visible=False)
    _save(fig, "09_pork_peers.png", "Company reports / results forecasts (Muyuan H1-26 = guidance midpoint)")

    # 10 Sow herd
    sh = DATA["sow_herd"]["rows"]
    fig, ax = plt.subplots(figsize=(10, 3.4))
    bars = ax.bar([r[0] for r in sh], [r[1] for r in sh], color=[MUTED, MUTED, MUTED, S1], **bar_kw)
    _bar_labels(ax, bars, "{:.1f}")
    ax.plot([r[0] for r in sh], [r[2] for r in sh], color=S2, lw=2, marker="o", ms=5, label="Official 'normal' level")
    ax.set_ylim(0, 48)
    ax.legend(loc="upper right")
    ax.grid(axis="x", visible=False)
    ax.set_title("China breeding sows, million head - capacity being cut, -14% from peak")
    _save(fig, "10_china_sow_herd.png", "NBS / MARA as quoted in COFCO Joycome FY2025 & H1 2026 results")

    # 11 Mengniu vs Yili
    d = DATA["dairy_peers"]["rows"]
    yrs = [str(r[0]) for r in d]
    fig, (a, b) = plt.subplots(1, 2, figsize=(11, 3.8))
    a.plot(yrs, [r[1] for r in d], color=S1, lw=2, marker="o", ms=5, label="Mengniu")
    a.plot(yrs, [r[2] for r in d], color=S2, lw=2, marker="o", ms=5, label="Yili")
    a.set_title("Revenue, RMB bn")
    a.legend(loc="upper left")
    a.set_ylim(60, 135)
    b.plot(yrs, [r[3] for r in d], color=S1, lw=2, marker="o", ms=5, label="Mengniu")
    b.plot(yrs, [r[4] for r in d], color=S2, lw=2, marker="o", ms=5, label="Yili")
    b.set_title("Attributable net profit, RMB bn")
    b.legend(loc="upper left")
    b.set_ylim(0, 13)
    for x_, v in zip(yrs, [r[3] for r in d]):
        if x_ in ("2024", "2025"):
            b.annotate(f"{v:.2f}", (x_, v), xytext=(0, -13), textcoords="offset points", ha="center", fontsize=8, color=INK2)
    _save(fig, "11_mengniu_vs_yili.png", "Company annual reports; Mengniu FY2025 presentation; Yili FY2025 (CLS)")

    # 12 Mengniu reported vs adjusted profit + segments
    fig, (a, b) = plt.subplots(1, 2, figsize=(11, 3.8))
    x = range(2)
    a.bar([i - 0.2 for i in x], [0.105, 1.545], 0.4, color=MUTED, label="Reported", **bar_kw)
    a.bar([i + 0.2 for i in x], [4.435, 3.960], 0.4, color=S1, label="Adjusted (ex-impairments)", **bar_kw)
    for i, (r1, r2) in enumerate([(0.105, 4.435), (1.545, 3.960)]):
        a.annotate(f"{r1:.2f}", (i - 0.2, r1), xytext=(0, 3), textcoords="offset points", ha="center", fontsize=8.5, color=INK2)
        a.annotate(f"{r2:.2f}", (i + 0.2, r2), xytext=(0, 3), textcoords="offset points", ha="center", fontsize=8.5, color=INK2)
    a.set_xticks(list(x), ["2024", "2025"])
    a.set_title("Mengniu net profit, RMB bn: reported vs adjusted")
    a.legend(loc="upper right", fontsize=8)
    a.set_ylim(0, 5.5)
    ms = DATA["mengniu_segments"]["rows"]
    g25 = [100 * (r[2] / r[1] - 1) for r in ms]
    g26 = [100 * (r[4] / r[3] - 1) for r in ms]
    y = range(len(ms))
    b.barh([i - 0.2 for i in y], g25, 0.4, color=MUTED, label="FY2025 YoY", **bar_kw)
    b.barh([i + 0.2 for i in y], g26, 0.4, color=S1, label="H1 2026 YoY", **bar_kw)
    for i, (u, v) in enumerate(zip(g25, g26)):
        b.annotate(f"{u:+.0f}%", (u, i - 0.2), xytext=(4 if u >= 0 else -4, 0), textcoords="offset points",
                   va="center", ha="left" if u >= 0 else "right", fontsize=8, color=INK2)
        b.annotate(f"{v:+.0f}%", (v, i + 0.2), xytext=(4, 0), textcoords="offset points", va="center", fontsize=8, color=INK2)
    b.axvline(0, color=INK2, lw=0.8)
    b.set_yticks(list(y), [r[0] for r in ms])
    b.invert_yaxis()
    b.set_xlim(-20, 42)
    b.set_title("Segment revenue growth, %")
    b.legend(loc="lower right", fontsize=8)
    b.grid(axis="y", visible=False)
    _save(fig, "12_mengniu_profit_segments.png", "Mengniu FY2025 and H1 2026 results presentations")

    # 13 Cycle map (diverging, analyst judgement)
    fig, ax = plt.subplots(figsize=(10, 2.6))
    cmap = LinearSegmentedColormap.from_list("div", ["#c43a3a", "#f1b8b0", "#ecebe7", "#a9c8ee", "#1f5fae"])
    mat = list(CYCLE["rows"].values())
    ax.imshow(mat, cmap=cmap, vmin=-2, vmax=2, aspect="auto")
    lab = {-2: "Deep trough", -1: "Weak", 0: "Neutral", 1: "Good", 2: "Boom"}
    for i, row in enumerate(mat):
        for k, v in enumerate(row):
            ax.text(k, i, lab[v], ha="center", va="center", fontsize=7.5, color=INK if abs(v) < 2 else "white")
    ax.set_xticks(range(len(CYCLE["years"])), CYCLE["years"])
    ax.set_yticks(range(3), list(CYCLE["rows"].keys()))
    ax.grid(False)
    for sp in ax.spines.values():
        sp.set_visible(False)
    ax.set_title("Earnings-cycle map: 2025 was the first year all three engines were weak at once")
    _save(fig, "13_cycle_map.png", "Analyst assessment based on the data in this pack (2027E = our forecast)")


# ---------------------------------------------------------------------------
# Workbook
# ---------------------------------------------------------------------------
def workbook():
    wb = Workbook()
    ws = wb.active
    ws.title = "README"
    ws["A1"] = "COFCO Corporation - analysis data pack"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A3"] = "Each sheet = one table used in the report. Sources: see the 'source' row and resources/SOURCES.md."
    ws["A4"] = "Prepared 13 Oct 2026. Educational analysis; not investment advice."
    hdr = PatternFill("solid", fgColor="1F3A5F")
    for key, t in DATA.items():
        sh = wb.create_sheet(key[:31])
        sh["A1"] = t["title"]
        sh["A1"].font = Font(bold=True, size=12)
        for j, c in enumerate(t["cols"], 1):
            cell = sh.cell(row=3, column=j, value=c)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = hdr
            cell.alignment = Alignment(wrap_text=True)
            sh.column_dimensions[get_column_letter(j)].width = 16 if j > 1 else 34
        for i, r in enumerate(t["rows"], 4):
            for j, v in enumerate(r, 1):
                sh.cell(row=i, column=j, value=v)
        n = 4 + len(t["rows"]) + 1
        sh.cell(row=n, column=1, value="Source tags: " + ", ".join(t["src"])).font = Font(italic=True)
        sh.cell(row=n + 1, column=1, value="Note: " + t["note"]).font = Font(italic=True)
    # computed ratios with live formulas
    sh = wb.create_sheet("ratios")
    sh["A1"] = "Computed ratios (live formulas)"
    sh["A1"].font = Font(bold=True, size=12)
    sh.append([])
    sh.append(["Grain traders net margin %", "2021", "2022", "2023", "2024", "2025"])
    for i in range(4):
        r = 4 + i
        sh.append([f"=traders!A{r}"] + [f"=traders!{get_column_letter(7 + k)}{r}/traders!{get_column_letter(2 + k)}{r}*100"
                                        for k in range(5)])
    sh.append([])
    sh.append(["COFCO group pre-tax margin %", "2021", "2022", "2023", "2024", "2025"])
    sh.append(["COFCO"] + [f"=group_financials!C{4 + k}/group_financials!B{4 + k}*100" for k in range(5)])
    sh.append([])
    sh.append(["Joycome profit before BA margin %"] + [r[0] for r in DATA["joycome"]["rows"]])
    sh.append(["Joycome"] + [f"=joycome!C{4 + k}/joycome!B{4 + k}*100" for k in range(len(DATA["joycome"]["rows"]))])
    sh.column_dimensions["A"].width = 34
    sh = wb.create_sheet("cycle_scores")
    sh.append(["sector"] + CYCLE["years"])
    for k, v in CYCLE["rows"].items():
        sh.append([k] + v)
    sh.append([])
    sh.append(["Scale: -2 deep trough, -1 weak, 0 neutral, +1 good, +2 boom. Analyst judgement."])
    wb.save(os.path.join(ROOT, "COFCO_Data_Pack.xlsx"))


if __name__ == "__main__":
    write_csvs()
    charts()
    workbook()
    print("built")
