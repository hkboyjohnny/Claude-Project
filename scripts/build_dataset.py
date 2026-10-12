"""Assemble the standardised peer dataset (data/dataset.json).

Programmatic sources : Eastmoney F10 (A-shares / HK / US balance+income), SEC XBRL (US cash flow, debt, D&A)
Hand-keyed from filings (each with a source note): Sinochem Energy, Syngenta AG, Pirelli, Michelin, Bridgestone, Nutrien
All amounts are millions of reporting currency, fiscal year ended 31 Dec unless noted.
"""
import json
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import extract_listed as x  # noqa: E402

ROOT = pathlib.Path(__file__).resolve().parent.parent
FIELDS = ["revenue", "cogs", "gross_profit", "ebit", "da", "ebitda", "interest_exp", "pretax", "net_income", "net_income_parent",
          "total_assets", "total_liabilities", "total_equity", "parent_equity", "cash", "current_assets", "current_liabilities",
          "debt", "ocf", "capex", "tax_surcharge"]

# IRS yearly-average rates for 2025, foreign currency per 1 USD (https://www.irs.gov/individuals/international-taxpayers/yearly-average-currency-exchange-rates)
FX = {"USD": 1.0, "CNY": 7.129, "EUR": 0.886, "JPY": 149.632}
FX_2024 = {"USD": 1.0, "CNY": 7.189, "EUR": 0.924, "JPY": 151.353}


def fill(d):
    """Derive gross profit / EBITDA where missing and make sure every field exists."""
    d = dict(d)
    if d.get("gross_profit") is None and d.get("cogs") is not None:
        d["gross_profit"] = d["revenue"] - d["cogs"]
    if d.get("ebitda") is None and d.get("ebit") is not None and d.get("da") is not None:
        d["ebitda"] = d["ebit"] + d["da"]
    for f in FIELDS:
        d.setdefault(f, None)
    return d


# ----------------------------------------------------------------------------------------------------------------------
# Hand-keyed filings
# ----------------------------------------------------------------------------------------------------------------------
SINOCHEM_ENERGY = {  # CNY m. Source: 中化能源股份有限公司 2025 consolidated financial statements (天职业字[2026]19618号) - statements pp.3-6, note 63
    2025: dict(revenue=394110.34, cogs=387812.60, ebit=-965.89 + 549.72, da=4308.95 + 581.46 + 313.54 + 210.55, interest_exp=1086.33,
               pretax=-937.97, net_income=-1682.75, net_income_parent=-1730.96,
               total_assets=105437.29, total_liabilities=68503.08, total_equity=36934.21, parent_equity=35684.15, cash=5870.23,
               current_assets=51155.55, current_liabilities=56071.97,
               debt=9675.83 + 6746.81 + 5213.45 + 2024.72 + 4467.38, ocf=6681.61, capex=1526.94, tax_surcharge=3601.87),
    2024: dict(revenue=426902.93, cogs=418942.94, ebit=-2746.64 + 527.18, da=4163.84 + 654.09 + 338.30 + 256.996, interest_exp=1127.09,
               pretax=-2923.94, net_income=-3182.19, net_income_parent=-3265.62,
               total_assets=102314.80, total_liabilities=65231.89, total_equity=37082.92, parent_equity=35817.85, cash=3431.37,
               current_assets=47001.62, current_liabilities=52991.94,
               debt=5758.31 + 9063.38 + 8825.76 + 0.0 + 2606.76, ocf=4793.61, capex=2433.08, tax_surcharge=5582.96),
}

SYNGENTA_AG = {  # USD m. Source: Syngenta AG Financial Report 2025 (IFRS) - consolidated statements pp.14-17, segment note
    2025: dict(revenue=16999, cogs=10281, ebit=1229, da=525 + 353, interest_exp=774, pretax=613, net_income=-152, net_income_parent=-156,
               total_assets=30799, total_liabilities=23412, total_equity=7387, parent_equity=7352, cash=677,
               current_assets=14922, current_liabilities=14082, debt=3581 + 6950, ocf=1526, capex=527 + 690),
    2024: dict(revenue=16981, cogs=10887, ebit=1431, da=566 + 295, interest_exp=903, pretax=694, net_income=261, net_income_parent=269,
               total_assets=29445, total_liabilities=22184, total_equity=7261, parent_equity=7246, cash=1025,
               current_assets=15100, current_liabilities=12208, debt=3039 + 7753, ocf=2571, capex=623 + 741),
}

PIRELLI = {  # EUR m. Source: Pirelli 2025 Annual Report consolidated statements (pp.544-548, cash flow) and FY2025 results release
    2025: dict(revenue=6776.157, ebit=891.203, da=1464.8 - 891.2, ebitda=1464.8, interest_exp=290.605, pretax=760.162, net_income=530.696,
               net_income_parent=497.527, total_assets=13353.183, total_liabilities=13353.183 - 6456.737, total_equity=6456.737,
               parent_equity=6277.848, cash=1525.886, current_assets=4134.659, current_liabilities=2854.387, debt=2747.192 + 210.602,
               ocf=1305.805, capex=352.286 + 25.580),
    2024: dict(revenue=6773.324, ebit=902.990, da=1475.7 - 903.0, ebitda=1475.7, interest_exp=422.911, pretax=647.793, net_income=501.104,
               net_income_parent=467.981, total_assets=13680.622, total_liabilities=13680.622 - 5912.254, total_equity=5912.254,
               parent_equity=5756.071, cash=1502.741, current_assets=4262.650, current_liabilities=3340.409, debt=3068.598 + 760.857,
               ocf=1292.564, capex=376.076 + 27.597),
}

MICHELIN = {  # EUR m. Source: Michelin 2025 Universal Registration Document, consolidated statements (sect. 5.2.1) and note 9.1
    2025: dict(revenue=25992, cogs=18973, ebit=2366, da=1943, interest_exp=207 + 53, pretax=2258, net_income=1664, net_income_parent=1665,
               total_assets=35018, total_liabilities=35018 - 18076, total_equity=18076, parent_equity=18072, cash=3877,
               current_assets=14271, current_liabilities=7503, debt=5050 + 726 + 612 + 257, ocf=3819, capex=2017),
    2024: dict(revenue=27193, cogs=19445, ebit=2631, da=1983, interest_exp=169 + 49, pretax=2445, net_income=1890, net_income_parent=1884,
               total_assets=37352, total_liabilities=37352 - 18634, total_equity=18634, parent_equity=18625, cash=3936,
               current_assets=15339, current_liabilities=8710, debt=4934 + 872 + 1375 + 258, ocf=4336, capex=2250),
}

BRIDGESTONE = {  # JPY m. Source: Bridgestone Consolidated Financial Results for FY2025 (IFRS) - statements pp.11-17
    2025: dict(revenue=4429452, cogs=2722789, ebit=381237, da=353229, interest_exp=44019, pretax=354661, net_income=334299,
               net_income_parent=327264, total_assets=5747705, total_liabilities=2027817, total_equity=3719888, parent_equity=3661793,
               cash=713810, current_assets=2863182, current_liabilities=1122654, debt=94958 + 392312 + 68871 + 270874, ocf=660442,
               capex=251091 + 36064),
    2024: dict(revenue=4430096, cogs=2704093, ebit=443319, da=348058, interest_exp=49802, pretax=421437, net_income=292181,
               net_income_parent=284989, total_assets=5723517, total_liabilities=1937029, total_equity=3786488, parent_equity=3731606,
               cash=706732, current_assets=2863632, current_liabilities=1176243, debt=138940 + 239441 + 66513 + 282827, ocf=548844,
               capex=299292 + 38001),
}

NUTRIEN = {  # USD m. Source: Nutrien 2025 audited consolidated financial statements (40-F exhibit 99.3); current lease liabilities from Eastmoney
    2025: dict(revenue=26885, cogs=936 + 17602, ebit=3049 + 687, da=2369, interest_exp=687, pretax=3049, net_income=2297, net_income_parent=2267,
               total_assets=52301, total_liabilities=26936, total_equity=25365, parent_equity=25323, cash=701, current_assets=14749,
               current_liabilities=11041, debt=873 + 513 + 9350 + 937 + 346, ocf=4007, capex=2005),
    2024: dict(revenue=25972, cogs=956 + 17486, ebit=1136 + 720, da=2339, interest_exp=720, pretax=1136, net_income=700, net_income_parent=674,
               total_assets=51840, total_liabilities=27398, total_equity=24442, parent_equity=24407, cash=853, current_assets=13792,
               current_liabilities=12045, debt=1534 + 1037 + 8881 + 999 + 356, ocf=3535, capex=2154),
}


# ----------------------------------------------------------------------------------------------------------------------
# Programmatic extractions for US (SEC + Eastmoney) and HK (Eastmoney)
# ----------------------------------------------------------------------------------------------------------------------
def us_peer(co, em, year, ebit_mode, debt_parts):
    """US-listed peer: SEC XBRL for cash flow / D&A / debt, Eastmoney for income lines not in the filtered SEC set."""
    inc, bal = x._em_long(em, "income", year), x._em_long(em, "balance", year)
    s = lambda tag: x.sec_val(co, tag, year)  # noqa: E731
    revenue = (s("Revenues") or s("RevenueFromContractWithCustomerExcludingAssessedTax")) / 1e6
    cogs = (s("CostOfGoodsAndServicesSold") or inc["营业成本"]) / 1e6
    pretax = (s("IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest")
              or s("IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments")) / 1e6
    interest = (s("InterestExpense") if co == "Mosaic" else -inc["利息支出"]) / 1e6
    op_inc = inc["营业利润"] if co == "Goodyear" else s("OperatingIncomeLoss")
    ebit = (pretax + interest) if ebit_mode == "pretax_plus_interest" else op_inc / 1e6
    ocf = (s("NetCashProvidedByUsedInOperatingActivities") or s("NetCashProvidedByUsedInOperatingActivitiesContinuingOperations")) / 1e6
    capex = (s("PaymentsToAcquirePropertyPlantAndEquipment") or s("PaymentsToAcquireProductiveAssets")) / 1e6
    debt = sum((s(t) if t.startswith(("Debt", "Long", "Short")) else bal.get(t, 0.0)) for t in debt_parts) / 1e6
    return dict(revenue=revenue, cogs=cogs, ebit=ebit, da=s("DepreciationDepletionAndAmortization") / 1e6, interest_exp=interest,
                pretax=pretax, net_income=(s("ProfitLoss") or inc["净利润"]) / 1e6, net_income_parent=s("NetIncomeLoss") / 1e6,
                total_assets=s("Assets") / 1e6, total_liabilities=bal["总负债"] / 1e6,
                total_equity=bal["股东权益合计"] / 1e6, parent_equity=s("StockholdersEquity") / 1e6, cash=s("CashAndCashEquivalentsAtCarryingValue") / 1e6,
                current_assets=s("AssetsCurrent") / 1e6, current_liabilities=s("LiabilitiesCurrent") / 1e6, debt=debt, ocf=ocf, capex=capex)


def hk_sinofert(year):
    inc = {r["STD_ITEM_CODE"]: float(r["AMOUNT"]) for r in json.loads((x.RAW / "em_00297.HK_income.json").read_text())
           if r["REPORT_DATE"][:10] == f"{year}-12-31" and r["DATE_TYPE_CODE"] == "001"}
    bal = {r["STD_ITEM_CODE"]: float(r["AMOUNT"]) for r in json.loads((x.RAW / "em_00297.HK_balance.json").read_text())
           if r["REPORT_DATE"][:10] == f"{year}-12-31" and r["DATE_TYPE_CODE"] == "001"}
    cf = {r["STD_ITEM_CODE"]: float(r["AMOUNT"]) for r in json.loads((x.RAW / "em_00297.HK_cashflow.json").read_text())
          if r["REPORT_DATE"][:10] == f"{year}-12-31" and r["DATE_TYPE_CODE"] == "001"}
    M = 1e6
    return dict(revenue=inc["004001001"] / M, cogs=inc["004005002"] / M, ebit=inc["004010999"] / M, da=cf["001009"] / M,
                interest_exp=inc["004011201"] / M, pretax=inc["004011999"] / M, net_income=inc["004012999"] / M,
                net_income_parent=inc["004025002"] / M, total_assets=bal["004009999"] / M, total_liabilities=bal["004025999"] / M,
                total_equity=bal["004036999"] / M, parent_equity=bal["004030999"] / M, cash=bal["004002010"] / M,
                current_assets=bal["004002999"] / M, current_liabilities=bal["004011999"] / M,
                debt=(bal["004011010"] + bal["004020001"] + bal["004011006"] + bal["004020005"]) / M,
                ocf=cf["003999"] / M, capex=(cf["005005"] + cf["005007"]) / M)


# ----------------------------------------------------------------------------------------------------------------------
# Entity catalogue
# ----------------------------------------------------------------------------------------------------------------------
# role: "Sinochem" = Sinochem Holdings group entity (anchor/related); "Peer" = comparison peer
E = []  # list of dicts


def add(key, name, sector, sub, role, ccy, listing, data, source, note=""):
    E.append(dict(key=key, name=name, sector=sector, sub=sub, role=role, ccy=ccy, listing=listing, source=source, note=note,
                  fy={str(y): fill(v) for y, v in data.items()}))


def a(code):
    return {2025: x.a_share(code, 2025), 2024: x.a_share(code, 2024)}


EM = "Eastmoney F10 (filed annual reports)"
SEC = "SEC EDGAR XBRL + Eastmoney US"

# ---- ENERGY --------------------------------------------------------------------------------------------------------
add("SCE", "Sinochem Energy (中化能源)", "Energy", "Trading, refining & marketing", "Sinochem", "CNY", "Unlisted (bond issuer)", SINOCHEM_ENERGY,
    "Audited FY2025 consolidated FS (天职国际), SSE bond disclosure", "Operating arm of Sinochem Corp's petrochemical segment; 15 mt/y refining capacity centred on Quanzhou")
add("SNP", "Sinopec Corp", "Energy", "Integrated major", "Peer", "CNY", "600028.SH / 0386.HK", a("600028.SH"), EM)
add("PTR", "PetroChina", "Energy", "Integrated major", "Peer", "CNY", "601857.SH / 0857.HK", a("601857.SH"), EM)
add("HLP", "Hengli Petrochemical", "Energy", "Private integrated refining-chemicals", "Peer", "CNY", "600346.SH", a("600346.SH"), EM)
add("RSP", "Rongsheng Petrochemical", "Energy", "Private integrated refining-chemicals", "Peer", "CNY", "002493.SZ", a("002493.SZ"), EM)
add("SSP", "Sinopec Shanghai Petrochemical", "Energy", "Refining & petrochemicals", "Peer", "CNY", "600688.SH / 0338.HK", a("600688.SH"), EM)

# ---- AGRICULTURE -----------------------------------------------------------------------------------------------------
add("SYN", "Syngenta AG (global CP + Seeds)", "Agriculture", "Crop protection & seeds", "Sinochem", "USD", "Unlisted (Syngenta Group sub.)", SYNGENTA_AG,
    "Syngenta AG Financial Report 2025 (IFRS)", "Excludes ADAMA, Syngenta Group China, Yangnong, Sinofert; includes $792m litigation provisions in 2025")
add("ADM", "ADAMA", "Agriculture", "Crop protection (generics)", "Sinochem", "CNY", "000553.SZ", a("000553.SZ"), EM, "Syngenta Group sub.")
add("YNG", "Yangnong Chemical", "Agriculture", "Crop protection (AI & intermediates)", "Sinochem", "CNY", "600486.SH", a("600486.SH"), EM, "Syngenta Group sub.")
add("SIF", "Sinofert", "Agriculture", "Fertilizer distribution", "Sinochem", "CNY", "0297.HK", {2025: hk_sinofert(2025), 2024: hk_sinofert(2024)},
    "Eastmoney HK F10 (filed annual reports)", "Spun out of Syngenta Group to Sinochem HK on 23 Dec 2025")
add("CTV", "Corteva", "Agriculture", "Crop protection & seeds", "Peer", "USD", "CTVA (NYSE)",
    {y: us_peer("Corteva", "CTVA.N", y, "pretax_plus_interest", ["DebtCurrent", "LongTermDebtAndCapitalLeaseObligations"]) for y in (2025, 2024)}, SEC,
    "Pre-separation consolidated FY2025")
add("FMC", "FMC Corp", "Agriculture", "Crop protection", "Peer", "USD", "FMC (NYSE)",
    {y: us_peer("FMC", "FMC.N", y, "operating", ["DebtCurrent", "LongTermDebtNoncurrent"]) for y in (2025, 2024)}, SEC,
    "FY2025 net loss of $2.2bn")
add("JSA", "Jiangshan Agrochemical", "Agriculture", "Crop protection (AI)", "Peer", "CNY", "600389.SH", a("600389.SH"), EM)
add("LCH", "Lier Chemical", "Agriculture", "Crop protection (AI & formulations)", "Peer", "CNY", "002258.SZ", a("002258.SZ"), EM)
add("NTR", "Nutrien", "Agriculture", "Crop nutrition & ag retail", "Peer", "USD", "NTR (NYSE/TSX)", NUTRIEN, "Nutrien FY2025 audited FS (40-F)")
add("MOS", "Mosaic", "Agriculture", "Crop nutrition", "Peer", "USD", "MOS (NYSE)",
    {y: us_peer("Mosaic", "MOS.N", y, "operating", ["ShortTermBorrowings", "LongTermDebtAndCapitalLeaseObligations", "长期负债(本期部分)"]) for y in (2025, 2024)}, SEC)
add("YTH", "Yuntianhua", "Agriculture", "Phosphate fertilizer", "Peer", "CNY", "600096.SH", a("600096.SH"), EM)
add("SLP", "Salt Lake Potash (青海盐湖)", "Agriculture", "Potash & lithium", "Peer", "CNY", "000792.SZ", a("000792.SZ"), EM,
    "Sinochem Corp holds a 5.73% stake (rating report)")

# ---- RUBBER ---------------------------------------------------------------------------------------------------------
add("PIR", "Pirelli", "Rubber", "Premium / High Value tyres", "Sinochem", "EUR", "PIRC.MI", PIRELLI, "Pirelli 2025 Annual Report (IFRS)",
    "Sinochem (via Marco Polo International Italy) holds ~34.1% (Dec 2025); Pirelli's accounts state no entity controls it under IFRS 10 since Apr 2025; no COGS line (by-nature P&L)")
add("AEO", "Aeolus Tyre (风神股份)", "Rubber", "Truck & bus radial tyres", "Sinochem", "CNY", "600469.SH", a("600469.SH"), EM, "ChemChina / CNTR sub.")
add("MIC", "Michelin", "Rubber", "Global tyre major", "Peer", "EUR", "ML.PA", MICHELIN, "Michelin 2025 URD (IFRS)")
add("BRI", "Bridgestone", "Rubber", "Global tyre major", "Peer", "JPY", "5108.T", BRIDGESTONE, "Bridgestone FY2025 results (IFRS)")
add("GYR", "Goodyear", "Rubber", "Global tyre major", "Peer", "USD", "GT (Nasdaq)",
    {y: us_peer("Goodyear", "GT.O", y, "operating", ["DebtCurrent", "LongTermDebtAndCapitalLeaseObligations"]) for y in (2025, 2024)}, SEC,
    "FY2025 net loss of $1.7bn")
add("LLT", "Linglong Tire", "Rubber", "China tyre producer", "Peer", "CNY", "601966.SH", a("601966.SH"), EM)
add("SLG", "Sailun Group", "Rubber", "China tyre producer", "Peer", "CNY", "601058.SH", a("601058.SH"), EM)
add("ZCR", "Zhongce Rubber", "Rubber", "China tyre producer", "Peer", "CNY", "603049.SH", a("603049.SH"), EM)
add("HNR", "Hainan Rubber", "Rubber", "Natural rubber (owns Halcyon Agri stake)", "Peer", "CNY", "601118.SH", a("601118.SH"), EM,
    "Agreed in Nov 2022 to buy 36% of Halcyon Agri from Sinochem International (completion not verified)")

# ---- H1 2026 vs H1 2025 interim snapshot (A-shares + Sinofert) -------------------------------------------------------------------
INTERIM = {}
for key, code in {"SNP": "600028.SH", "PTR": "601857.SH", "HLP": "600346.SH", "RSP": "002493.SZ", "SSP": "600688.SH", "ADM": "000553.SZ",
                  "YNG": "600486.SH", "JSA": "600389.SH", "LCH": "002258.SZ", "YTH": "600096.SH", "SLP": "000792.SZ", "AEO": "600469.SH",
                  "LLT": "601966.SH", "SLG": "601058.SH", "ZCR": "603049.SH", "HNR": "601118.SH"}.items():
    try:
        INTERIM[key] = {"h1_2026": x.a_share_interim(code, "2026-06-30"), "h1_2025": x.a_share_interim(code, "2025-06-30")}
    except StopIteration:
        pass

# Sinofert (HK) interim: cumulative six-month income lines (type 002)
def _hk_interim(d):
    rows = {r["STD_ITEM_CODE"]: float(r["AMOUNT"]) for r in json.loads((x.RAW / "em_00297.HK_income.json").read_text())
            if r["REPORT_DATE"][:10] == d and r["DATE_TYPE_CODE"] == "002"}
    return dict(revenue=rows["004001001"] / 1e6, net_income=rows["004012999"] / 1e6, net_income_parent=rows["004025002"] / 1e6,
                pretax=rows["004011999"] / 1e6)

INTERIM["SIF"] = {"h1_2026": _hk_interim("2026-06-30"), "h1_2025": _hk_interim("2025-06-30")}

if __name__ == "__main__":
    out = dict(fx_2025=FX, fx_2024=FX_2024, fields=FIELDS, entities=E, interim=INTERIM)
    (ROOT / "data" / "dataset.json").write_text(json.dumps(out, ensure_ascii=False, indent=1))
    print("entities", len(E), "interim", len(INTERIM))
    for e in E:
        f = e["fy"]["2025"]
        print(f"{e['key']:4} {e['name'][:34]:34} {e['ccy']} rev {f['revenue']:>11,.0f} ebitda {f['ebitda']:>9,.0f} ni {f['net_income']:>9,.0f} debt {f['debt']:>9,.0f} TA {f['total_assets']:>10,.0f}")
