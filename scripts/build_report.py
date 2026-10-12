"""Fill the report template with tables generated from the shipped workbook."""
import pathlib, sys
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import report_tables as rt

ROOT = pathlib.Path(__file__).resolve().parent.parent
t = (ROOT / "report_src" / "report_template.md").read_text()
for sec in ("Energy", "Agriculture", "Rubber"):
    t = t.replace(f"<!--TABLE:{sec}-->", rt.table(sec, rt.STATS[sec]))
t = t.replace("<!--SCORECARD-->", rt.scorecard_summary())
(ROOT / "Sinochem_Peer_Comparison_Report.md").write_text(t)
print("report written", len(t.split()), "words")
