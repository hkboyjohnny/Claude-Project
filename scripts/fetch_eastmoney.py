"""Fetch full financial statements for listed peers from Eastmoney's public F10 datasets.

A-shares : RPT_F10_FINANCE_{GINCOME,GBALANCE,GCASHFLOW}  (one row per period, wide format)
HK / US  : RPT_HKF10_FN_* / RPT_USF10_FN_*              (one row per line item, long format)

Raw JSON is written to data/raw/ so every figure in the workbook is traceable.
"""
import json, sys, time, urllib.parse, urllib.request, pathlib

OUT = pathlib.Path(__file__).resolve().parent.parent / "data" / "raw"
OUT.mkdir(parents=True, exist_ok=True)
BASE = "https://datacenter.eastmoney.com/securities/api/data/v1/get"
UA = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"

A_SHARES = {  # secucode -> label
    "600028.SH": "Sinopec Corp", "601857.SH": "PetroChina", "600346.SH": "Hengli Petrochemical",
    "002493.SZ": "Rongsheng Petrochemical", "600688.SH": "Sinopec Shanghai Petrochemical",
    "000553.SZ": "ADAMA", "600096.SH": "Yuntianhua", "000792.SZ": "Salt Lake Potash",
    "600469.SH": "Aeolus Tyre", "601966.SH": "Linglong Tire", "601058.SH": "Sailun Group",
    "603049.SH": "Zhongce Rubber", "601118.SH": "Hainan Rubber", "600500.SH": "Sinochem International",
    "600486.SH": "Yangnong Chemical", "600389.SH": "Jiangshan Agrochemical", "002258.SZ": "Lier Chemical",
}
HK = {"00297.HK": "Sinofert"}
US = {"CTVA.N": "Corteva", "FMC.N": "FMC", "NTR.N": "Nutrien", "MOS.N": "Mosaic", "GT.O": "Goodyear"}


def get(params, tries=4):
    url = BASE + "?" + urllib.parse.urlencode(params, safe="()")
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA})
            with urllib.request.urlopen(req, timeout=40) as r:
                return json.load(r)
        except Exception as e:  # noqa: BLE001
            if i == tries - 1:
                raise
            time.sleep(2 ** (i + 1))


def fetch(report, filt, size):
    p = {"reportName": report, "columns": "ALL", "filter": filt, "pageNumber": 1, "pageSize": size,
         "sortColumns": "REPORT_DATE", "sortTypes": -1, "source": "HSF10", "client": "PC"}
    d = get(p)
    if not d.get("success"):
        raise RuntimeError(f"{report} {filt}: {d.get('message')}")
    return d["result"]["data"]


def main():
    for code, name in (A_SHARES.items() if "--all" in sys.argv else []):
        for tag, rep in (("income", "RPT_F10_FINANCE_GINCOME"), ("balance", "RPT_F10_FINANCE_GBALANCE"),
                         ("cashflow", "RPT_F10_FINANCE_GCASHFLOW")):
            rows = fetch(rep, f'(SECUCODE="{code}")', 24)
            (OUT / f"em_{code}_{tag}.json").write_text(json.dumps(rows, ensure_ascii=False))
        print("A-share ok", code, name, len(rows))
    for code, name in HK.items():
        for tag, rep in (("income", "RPT_HKF10_FN_INCOME"), ("balance", "RPT_HKF10_FN_BALANCE"),
                         ("cashflow", "RPT_HKF10_FN_CASHFLOW_PC")):
            rows = fetch(rep, f'(SECUCODE="{code}")', 1500)
            (OUT / f"em_{code}_{tag}.json").write_text(json.dumps(rows, ensure_ascii=False))
        print("HK ok", code, name, len(rows))
    for code, name in US.items():
        # US cash flow is not exposed by this API; OCF/D&A/capex come from SEC XBRL (fetch_sec.py)
        for tag, rep in (("income", "RPT_USF10_FN_INCOME"), ("balance", "RPT_USF10_FN_BALANCE")):
            rows = fetch(rep, f'(SECUCODE="{code}")', 1500)
            (OUT / f"em_{code}_{tag}.json").write_text(json.dumps(rows, ensure_ascii=False))
        print("US ok", code, name, len(rows))


if __name__ == "__main__":
    sys.exit(main())
