"""Export the Proposal Log sheet to the JSON file the online dashboard reads.

Usage (from the repo root):
    python3 tools/export_proposals_json.py [workbook.xlsx] [output.json]

Reads every real entry on the "Proposal Log" sheet (row 6 down; row 5 is the
grey example row) and skips rows with no Job. Blank Yes/No cells count as "No",
the same way the Summary sheet treats them.
"""
import json
import sys
from datetime import date, datetime

from openpyxl import load_workbook

XLSX = sys.argv[1] if len(sys.argv) > 1 else "upwork-proposal-log.xlsx"
OUT = sys.argv[2] if len(sys.argv) > 2 else "proposals.json"

FIRST_ROW = 6
KEYS = ["date", "job", "connects", "bid", "viewed", "replied", "interviewed", "hired", "notes"]


def number(value):
    if value is None or value == "":
        return None
    value = float(value)
    return int(value) if value.is_integer() else value


def yes_no(value):
    return "Yes" if str(value or "").strip().lower() == "yes" else "No"


def iso_date(value):
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return str(value or "")


sheet = load_workbook(XLSX, data_only=True)["Proposal Log"]
rows = []
for cells in sheet.iter_rows(min_row=FIRST_ROW, max_col=len(KEYS), values_only=True):
    record = dict(zip(KEYS, cells))
    if not record["job"]:
        continue
    rows.append(
        {
            "date": iso_date(record["date"]),
            "job": str(record["job"]).strip(),
            "connects": number(record["connects"]),
            "bid": number(record["bid"]),
            "viewed": yes_no(record["viewed"]),
            "replied": yes_no(record["replied"]),
            "interviewed": yes_no(record["interviewed"]),
            "hired": yes_no(record["hired"]),
            "notes": str(record["notes"] or "").strip(),
        }
    )

with open(OUT, "w", encoding="utf-8") as f:
    json.dump(rows, f, indent=1, ensure_ascii=False)
print(f"wrote {len(rows)} proposal(s) to {OUT}")
