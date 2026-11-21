#!/usr/bin/env python3
"""
Builds a folder of MADE-UP travel-expense claim workbooks to try the extractor on.

    python make_sample_claims.py sample_claims

Every file has one sheet named "Traveling Expense" with the values in the cells that extract_expenses.py reads (see FIELD_MAP).
It also creates two files the extractor must NOT turn into rows: a workbook with the wrong sheet name (logged as an error)
and an Excel lock file (~$...), which is skipped.
"""
from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

COMPANY = "ABCD Pharma (Pvt.) Ltd"
SHEET = "Traveling Expense"

# (relative folder, file name, name, employee id, group, designation, month, town, km, dailies, ex, on, trav, sundries)
CLAIMS = [
    ("Group A/Lahore", "Claim_90001_Jun-2026.xlsx", "Imran Sample", 90001, "Group A", "Territory Manager", "June-2026", "Lahore", 1480, 22, 3, 1, 36550, 4200),
    ("Group A/Lahore", "Claim_90002_Jun-2026.xlsx", "Sana Example", 90002, "Group A", "Senior Territory Manager", "June-2026", "Lahore", 1725, 24, 4, 2, 42380, 5100),
    ("Group A/Karachi", "Claim_90003_Jun-2026.xlsx", "Bilal Demo", 90003, "Group A", "Territory Manager", "June-2026", "Karachi", 1310, 21, 2, 0, 31870, 3300),
    ("Group A/Karachi", "Claim_90004_Jun-2026.xlsx", "Nadia Test", 90004, "Group A", "Area Manager", "June-2026", "Karachi", 2050, 25, 5, 3, 51900, 6050),
    ("Group A/Karachi", "Claim_90005_Jun-2026.xlsx", "Farhan Sample", 90005, "Group A", "Territory Manager", "June-2026", "Hyderabad", 980, 18, 1, 0, 24100, 2250),
    ("Group B/Islamabad", "Claim_90006_Jun-2026.xlsx", "Ayesha Example", 90006, "Group B", "Regional Manager", "June-2026", "Islamabad", 2380, 26, 6, 4, 61450, 7800),
    ("Group B/Islamabad", "Claim_90007_Jun-2026.xlsx", "Hina Test", 90007, "Group B", "Territory Manager", "June-2026", "Rawalpindi", 1195, 20, 2, 1, 29240, 3050),
    ("Group B/Multan", "Claim_90008_Jun-2026.xlsx", "Kashif Demo", 90008, "Group B", "Senior Territory Manager", "June-2026", "Multan", 1860, 23, 3, 2, 45120, 4900),
    ("Group B/Multan", "Claim_90009_Jun-2026.xlsx", "Zara Sample", 90009, "Group B", "Sales Executive", "June-2026", "Bahawalpur", 760, 15, 0, 0, 18430, 1700),
    ("Group B", "Claim_90010_Jun-2026.xlsx", "Omar Example", 90010, "Group B", "Area Manager", "June-2026", "Faisalabad", 2210, 25, 4, 2, 55300, 6400),
    ("", "Claim_90011_Jun-2026.xlsx", "Maryam Test", 90011, "Head Office", "Territory Manager", "June-2026", "Peshawar", 1040, 19, 2, 1, 25760, 2850),
]


def write_claim(path: Path, rec: tuple, sheet: str = SHEET) -> None:
    _, _, name, emp_id, group, desig, month, town, km, dailies, ex, on, trav, sundries = rec
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    bold = Font(name="Arial", bold=True)
    ws["A1"] = f"{COMPANY}  -  FIELD FORCE TRAVEL EXPENSE CLAIM"
    ws["A1"].font = Font(name="Arial", bold=True, size=13)
    labels = {"A7": "Name", "G7": "Employee ID", "J7": "Group", "A9": "Designation", "G9": "Month", "J9": "Town",
              "A45": "TOTAL KM", "A59": "Total Number of Dailies", "A61": "Total Number of EX", "A63": "Total Number of ON",
              "G47": "Travelling Expense Total", "G48": "Sundries Total", "G49": "GRAND TOTAL"}
    for addr, text in labels.items():
        ws[addr] = text
        ws[addr].font = bold
    # the cells the extractor reads (FIELD_MAP)
    values = {"B7": name, "H7": emp_id, "K7": group, "C9": desig, "H9": month, "K9": town, "F45": km,
              "E59": dailies, "E61": ex, "E63": on, "I47": trav, "I48": sundries, "I49": trav + sundries}
    thin = Side(style="thin", color="999999")
    for addr, v in values.items():
        ws[addr] = v
        ws[addr].border = Border(left=thin, right=thin, top=thin, bottom=thin)
        ws[addr].fill = PatternFill("solid", start_color="EEF3FA")
        if isinstance(v, (int, float)) and addr[0] in "FI":
            ws[addr].number_format = "#,##0"
        ws[addr].alignment = Alignment(horizontal="left" if isinstance(v, str) else "right")
    ws.column_dimensions["A"].width = 26
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def main() -> int:
    out = Path(sys.argv[1] if len(sys.argv) > 1 else "sample_claims")
    for rec in CLAIMS:
        write_claim(out / rec[0] / rec[1], rec)
    # a claim saved with the wrong sheet name: the extractor must report it, not guess
    write_claim(out / "Group B/Multan" / "Claim_90099_WrongSheet.xlsx", CLAIMS[0], sheet="Sheet1")
    # an Excel lock file: ignored
    (out / "Group A/Lahore").mkdir(parents=True, exist_ok=True)
    (out / "Group A/Lahore" / "~$Claim_90001_Jun-2026.xlsx").write_bytes(b"lock")
    print(f"Wrote {len(CLAIMS) + 1} sample workbooks (+1 lock file) under {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
