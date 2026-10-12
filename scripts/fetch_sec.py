"""Pull annual (10-K / 40-F) XBRL facts from SEC EDGAR for the US-listed peers and keep only the tags we use."""
import json, pathlib, sys, urllib.request

OUT = pathlib.Path(__file__).resolve().parent.parent / "data" / "raw"
UA = "Peer-comparison research research@example.com"
CIKS = {"Corteva": 1755672, "FMC": 37785, "Mosaic": 1285785, "Goodyear": 42582, "Nutrien": 1725964}
TAGS = [
    "Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax", "SalesRevenueNet", "CostOfRevenue", "CostOfGoodsAndServicesSold",
    "GrossProfit", "OperatingIncomeLoss", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
    "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
    "NetIncomeLoss", "ProfitLoss", "InterestExpense", "InterestExpenseNonoperating", "InterestExpenseDebt",
    "DepreciationDepletionAndAmortization", "DepreciationAndAmortization", "DepreciationAmortizationAndAccretionNet",
    "NetCashProvidedByUsedInOperatingActivities", "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations", "PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireProductiveAssets",
    "PaymentsToAcquireIntangibleAssets", "PaymentsForSoftware", "Assets", "AssetsCurrent", "Liabilities", "LiabilitiesCurrent",
    "StockholdersEquity", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest",
    "CashAndCashEquivalentsAtCarryingValue", "LongTermDebt", "LongTermDebtNoncurrent", "LongTermDebtCurrent", "ShortTermBorrowings",
    "DebtCurrent", "ShortTermBankLoansAndNotesPayable", "NotesPayableCurrent", "LongTermDebtAndCapitalLeaseObligations",
    "FinanceLeaseLiabilityNoncurrent", "OperatingLeaseLiabilityNoncurrent", "InventoryNet", "LiabilitiesAndStockholdersEquity",
]


def main():
    for name, cik in CIKS.items():
        url = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json"
        req = urllib.request.Request(url, headers={"User-Agent": UA})
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                d = json.load(r)
        except Exception as e:  # noqa: BLE001
            print("FAIL", name, e)
            continue
        keep = {}
        for taxo, facts in d.get("facts", {}).items():
            if taxo == "dei":
                continue
            for tag, v in facts.items():
                if tag in TAGS:
                    units = {}
                    for unit, arr in v["units"].items():
                        arr = [a for a in arr if a.get("fp") == "FY" and a.get("form") in ("10-K", "40-F", "10-K/A") and a.get("fy", 0) >= 2024]
                        if arr:
                            units[unit] = arr
                    if units:
                        keep[f"{taxo}:{tag}"] = units
        (OUT / f"sec_{name}.json").write_text(json.dumps({"entity": d.get("entityName"), "facts": keep}))
        print("ok", name, d.get("entityName"), len(keep))


if __name__ == "__main__":
    sys.exit(main())
