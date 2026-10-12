"""Ratio definitions (mirrored by the Excel formulas) - used to cross-check the workbook after recalculation."""
import json, pathlib, statistics as st

ROOT = pathlib.Path(__file__).resolve().parent.parent
D = json.loads((ROOT / "data" / "dataset.json").read_text())


def ratios(e):
    c, p = e["fy"]["2025"], e["fy"]["2024"]
    fx = D["fx_2025"][e["ccy"]]
    div = lambda a, b: (a / b) if (a is not None and b not in (None, 0)) else None  # noqa: E731
    avg = lambda k: (c[k] + p[k]) / 2  # noqa: E731
    nd = c["debt"] - c["cash"]
    return {
        "Revenue (USD bn)": c["revenue"] / fx / 1000,
        "Total assets (USD bn)": c["total_assets"] / fx / 1000,
        "Revenue growth": div(c["revenue"] - p["revenue"], p["revenue"]),
        "EBITDA growth": div(c["ebitda"] - p["ebitda"], abs(p["ebitda"])),
        "Gross margin": div(c["gross_profit"], c["revenue"]),
        "EBITDA margin": div(c["ebitda"], c["revenue"]),
        "EBIT margin": div(c["ebit"], c["revenue"]),
        "Net margin": div(c["net_income"], c["revenue"]),
        "ROE": div(c["net_income"], avg("total_equity")),
        "ROA": div(c["net_income"], avg("total_assets")),
        "Asset turnover (x)": div(c["revenue"], avg("total_assets")),
        "Liabilities / assets": div(c["total_liabilities"], c["total_assets"]),
        "Debt / equity": div(c["debt"], c["total_equity"]),
        "Net debt / EBITDA (x)": div(nd, c["ebitda"]) if c["ebitda"] > 0 else None,
        "Interest cover (x)": div(c["ebit"], c["interest_exp"]),
        "Current ratio (x)": div(c["current_assets"], c["current_liabilities"]),
        "OCF / revenue": div(c["ocf"], c["revenue"]),
        "OCF / EBITDA": div(c["ocf"], c["ebitda"]) if c["ebitda"] > 0 else None,
        "Capex / revenue": div(c["capex"], c["revenue"]),
        "FCF margin": div(c["ocf"] - c["capex"], c["revenue"]),
    }


if __name__ == "__main__":
    import sys
    sec = sys.argv[1] if len(sys.argv) > 1 else None
    keys = None
    for e in D["entities"]:
        if sec and e["sector"] != sec:
            continue
        r = ratios(e)
        if keys is None:
            keys = list(r)
            print(f"{'':36}" + "".join(f"{k[:11]:>12}" for k in keys))
        def f(k, v):
            if v is None: return "n/a"
            if "USD" in k: return f"{v:,.1f}"
            if "(x)" in k: return f"{v:.2f}"
            return f"{100*v:.1f}%"
        print(f"{e['role'][:3]} {e['name'][:32]:32}" + "".join(f"{f(k, r[k]):>12}" for k in keys))
