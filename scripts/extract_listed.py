"""Standardise raw Eastmoney / SEC data into a common schema (millions of reporting currency).

Common fields per entity-year:
 revenue, cogs, gross_profit, ebit, da, ebitda, interest_exp, pretax, net_income, net_income_parent,
 total_assets, total_liabilities, total_equity, parent_equity, cash, current_assets, current_liabilities,
 debt, ocf, capex

Definitions (applied to every entity so the matrix is like-for-like):
 EBIT   = operating profit before financing as reported (China GAAP: operating profit + net finance expense;
          IFRS / US GAAP: reported operating income, or closest equivalent)
 EBITDA = EBIT + depreciation & amortisation (no add-back for impairments or one-offs)
 Debt   = interest-bearing borrowings + bonds + lease liabilities (balance-sheet basis)
 Capex  = cash paid for PP&E and intangibles
"""
import json, pathlib

RAW = pathlib.Path(__file__).resolve().parent.parent / "data" / "raw"
M = 1e6


def _load(name):
    return json.loads((RAW / name).read_text())


def _g(row, k):
    v = row.get(k)
    return 0.0 if v in (None, "") else float(v)


def a_share(secucode, year):
    """Annual full-year figures for an A-share (China GAAP) from Eastmoney F10."""
    d = f"{year}-12-31"
    pick = lambda tag: next(r for r in _load(f"em_{secucode}_{tag}.json") if r["REPORT_DATE"][:10] == d and r["REPORT_TYPE"] == "年报")  # noqa: E731
    i, b, c = pick("income"), pick("balance"), pick("cashflow")
    fa, og = _g(c, "FA_IR_DEPR"), _g(c, "OILGAS_BIOLOGY_DEPR")
    dep = fa if (og == fa or og == 0) else fa + og          # same combined line is mapped twice for some filers
    da = dep + _g(c, "IR_DEPR") + _g(c, "IA_AMORTIZE") + _g(c, "LPE_AMORTIZE") + _g(c, "USERIGHT_ASSET_AMORTIZE")
    revenue = _g(i, "TOTAL_OPERATE_INCOME")
    ebit = _g(i, "OPERATE_PROFIT") + _g(i, "FINANCE_EXPENSE")
    debt = sum(_g(b, k) for k in ("SHORT_LOAN", "LONG_LOAN", "BOND_PAYABLE", "NONCURRENT_LIAB_1YEAR", "LEASE_LIAB"))
    out = dict(
        revenue=revenue, cogs=_g(i, "OPERATE_COST"), ebit=ebit, da=da, interest_exp=_g(i, "FE_INTEREST_EXPENSE"),
        pretax=_g(i, "TOTAL_PROFIT"), net_income=_g(i, "NETPROFIT"), net_income_parent=_g(i, "PARENT_NETPROFIT"),
        total_assets=_g(b, "TOTAL_ASSETS"), total_liabilities=_g(b, "TOTAL_LIABILITIES"), total_equity=_g(b, "TOTAL_EQUITY"),
        parent_equity=_g(b, "TOTAL_PARENT_EQUITY"), cash=_g(b, "MONETARYFUNDS"),
        current_assets=_g(b, "TOTAL_CURRENT_ASSETS"), current_liabilities=_g(b, "TOTAL_CURRENT_LIAB"), debt=debt,
        ocf=_g(c, "NETCASH_OPERATE"), capex=_g(c, "CONSTRUCT_LONG_ASSET"), tax_surcharge=_g(i, "OPERATE_TAX_ADD"),
    )
    out = {k: v / M for k, v in out.items()}
    out["gross_profit"] = out["revenue"] - out["cogs"]
    out["ebitda"] = out["ebit"] + out["da"]
    out["_notice_date"] = i["NOTICE_DATE"][:10]
    return out


def a_share_interim(secucode, d):
    """Cumulative interim figures (e.g. H1) - revenue / net income only."""
    i = next(r for r in _load(f"em_{secucode}_income.json") if r["REPORT_DATE"][:10] == d)
    return dict(revenue=_g(i, "TOTAL_OPERATE_INCOME") / M, net_income=_g(i, "NETPROFIT") / M,
                net_income_parent=_g(i, "PARENT_NETPROFIT") / M, pretax=_g(i, "TOTAL_PROFIT") / M)


def _hk(secucode, tag, year, code_key):
    rows = [r for r in _load(f"em_{secucode}_{tag}.json") if r["REPORT_DATE"][:10] == f"{year}-12-31" and r["DATE_TYPE_CODE"] == "001"]
    return {r.get("STD_ITEM_NAME") or r.get("ITEM_NAME"): float(r["AMOUNT"]) for r in rows}


def _em_long(secucode, tag, year, dtype="001"):
    rows = [r for r in _load(f"em_{secucode}_{tag}.json") if r["REPORT_DATE"][:10] == f"{year}-12-31" and r["DATE_TYPE_CODE"] == dtype]
    return {r["ITEM_NAME"]: float(r["AMOUNT"]) for r in rows if r.get("AMOUNT") is not None}


def sec_val(co, tag, year, taxo=None):
    """Annual (duration > 300d or instant) value from filtered SEC XBRL facts for FY end <year>-12-31."""
    from datetime import date
    d = _load(f"sec_{co}.json")["facts"]
    for key, units in d.items():
        t, name = key.split(":")
        if name != tag or (taxo and t != taxo):
            continue
        for arr in units.values():
            vals = [a for a in arr if a["end"] == f"{year}-12-31" and ("start" not in a or (date.fromisoformat(a["end"]) - date.fromisoformat(a["start"])).days > 300)]
            if vals:
                return float(vals[-1]["val"])
    return None


if __name__ == "__main__":
    codes = ["600028.SH", "601857.SH", "600346.SH", "002493.SZ", "600688.SH", "000553.SZ", "600486.SH", "600096.SH", "000792.SZ",
             "600389.SH", "002258.SZ", "600469.SH", "601966.SH", "601058.SH", "603049.SH", "601118.SH", "600500.SH"]
    hdr = f"{'code':10} {'yr':4} {'rev':>11} {'GP%':>6} {'EBITDA':>9} {'EBIT':>9} {'DA':>8} {'NI':>9} {'NIpar':>9} {'TA':>11} {'Eq':>10} {'Debt':>9} {'Cash':>9} {'OCF':>9} {'Capex':>8}"
    print(hdr)
    for c in codes:
        for y in (2025, 2024):
            x = a_share(c, y)
            print(f"{c:10} {y} {x['revenue']:11.0f} {100*x['gross_profit']/x['revenue']:6.1f} {x['ebitda']:9.0f} {x['ebit']:9.0f} {x['da']:8.0f} {x['net_income']:9.0f} {x['net_income_parent']:9.0f} {x['total_assets']:11.0f} {x['total_equity']:10.0f} {x['debt']:9.0f} {x['cash']:9.0f} {x['ocf']:9.0f} {x['capex']:8.0f}")
