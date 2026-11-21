#!/usr/bin/env python3
"""
Field Force Travel Expense Claim — batch extractor (v2 with fallback reader).

Walks a folder tree, reads each .xlsx/.xlsm using fixed cell addresses,
and writes one consolidated .xlsx with one row per source file.

Reading strategy:
  1. Try openpyxl (strict, fast on clean files).
  2. If that fails with a malformed-XML error, try python-calamine
     (lenient Rust-based reader). This handles files saved by mobile
     Excel apps, WPS Office, LibreOffice, or older Excel versions whose
     XML doesn't perfectly match the spec.
  3. Only if both fail is the file logged as an error.

Source files are never modified.

Usage:
    python extract_expenses.py <root-input-folder> <output-xlsx-path>
"""
from __future__ import annotations

import sys
import warnings
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Optional fallback reader. The script still works if it isn't installed —
# files that openpyxl can't read will simply be logged as errors.
try:
    from python_calamine import CalamineWorkbook
    CALAMINE_AVAILABLE = True
except ImportError:
    CALAMINE_AVAILABLE = False


FIELD_MAP: dict[str, tuple[str, str]] = {
    "Name":                    ("Traveling Expense", "B7"),
    "Employee ID":             ("Traveling Expense", "H7"),
    "Group":                   ("Traveling Expense", "K7"),
    "Designation":             ("Traveling Expense", "C9"),
    "Month":                   ("Traveling Expense", "H9"),
    "Town":                    ("Traveling Expense", "K9"),
    "Total KM":                ("Traveling Expense", "F45"),
    "Total Number of Dailies": ("Traveling Expense", "E59"),
    "Total Number of EX":      ("Traveling Expense", "E61"),
    "Total Number of ON":      ("Traveling Expense", "E63"),
    "Trav. Exp. Total":        ("Traveling Expense", "I47"),
    "Sundries Total":          ("Traveling Expense", "I48"),
    "Grand Total":             ("Traveling Expense", "I49"),
}

OUTPUT_COLUMNS = [
    "Category",
    "Sub-Category",
    "Source File",
    "Source Path",
    "Read Method",  # which reader succeeded — useful for audit traceability
    *FIELD_MAP.keys(),
]


def derive_categories(rel_path: Path) -> tuple[str, str]:
    parts = rel_path.parts[:-1]
    if not parts:
        return "", ""
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], "/".join(parts[1:])


# ---------------------------------------------------------------------------
# Reading
# ---------------------------------------------------------------------------

def _addr_to_rc(addr: str) -> tuple[int, int]:
    """Convert an Excel cell address like 'B7' to 0-indexed (row, col)."""
    col_str = ''.join(c for c in addr if c.isalpha()).upper()
    row_num = int(''.join(c for c in addr if c.isdigit()))
    col_num = 0
    for ch in col_str:
        col_num = col_num * 26 + (ord(ch) - ord('A') + 1)
    return row_num - 1, col_num - 1


def _normalize(value):
    """Make calamine output match openpyxl conventions (e.g. int for whole numbers)."""
    if value == "" or value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def read_with_openpyxl(file_path: Path) -> dict[str, object]:
    wb = load_workbook(file_path, data_only=True)
    try:
        out: dict[str, object] = {}
        for col_name, (sheet_name, cell_ref) in FIELD_MAP.items():
            if sheet_name not in wb.sheetnames:
                raise ValueError(f"sheet '{sheet_name}' not found")
            out[col_name] = wb[sheet_name][cell_ref].value
        return out
    finally:
        wb.close()


def read_with_calamine(file_path: Path) -> dict[str, object]:
    wb = CalamineWorkbook.from_path(str(file_path))
    out: dict[str, object] = {}
    for col_name, (sheet_name, cell_ref) in FIELD_MAP.items():
        if sheet_name not in wb.sheet_names:
            raise ValueError(f"sheet '{sheet_name}' not found")
        ws_data = wb.get_sheet_by_name(sheet_name).to_python()
        r, c = _addr_to_rc(cell_ref)
        if r < len(ws_data) and c < len(ws_data[r]):
            out[col_name] = _normalize(ws_data[r][c])
        else:
            out[col_name] = None
    return out


def extract_one(file_path: Path) -> tuple[dict[str, object], str]:
    """
    Try openpyxl, then calamine. Returns (fields, method_used).
    Raises only if both fail.
    """
    try:
        return read_with_openpyxl(file_path), "openpyxl"
    except Exception as e_openpyxl:
        # Layout problems (wrong sheet) are NOT fixable by a different reader —
        # bail out immediately so the user sees the real reason.
        if "not found" in str(e_openpyxl):
            raise
        if not CALAMINE_AVAILABLE:
            raise RuntimeError(
                f"{type(e_openpyxl).__name__}: {e_openpyxl} "
                f"[install python-calamine to enable fallback reader]"
            )
        try:
            return read_with_calamine(file_path), "calamine"
        except Exception as e_calamine:
            raise RuntimeError(
                f"openpyxl: {type(e_openpyxl).__name__}: {e_openpyxl} | "
                f"calamine: {type(e_calamine).__name__}: {e_calamine}"
            )


# ---------------------------------------------------------------------------
# File walking and output
# ---------------------------------------------------------------------------

def find_excel_files(root: Path):
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        if path.suffix.lower() not in (".xlsx", ".xlsm"):
            continue
        if path.name.startswith("~$") or path.name.startswith("."):
            continue
        yield path


def write_output(rows: list[dict], output_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Consolidated"

    ws.append(OUTPUT_COLUMNS)
    header_font = Font(name="Arial", bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", start_color="305496")
    header_align = Alignment(horizontal="center", vertical="center")
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align

    body_font = Font(name="Arial")
    for row in rows:
        ws.append([row.get(col, "") for col in OUTPUT_COLUMNS])
    for row_cells in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row_cells:
            cell.font = body_font

    ws.freeze_panes = "A2"

    for col_idx, col_name in enumerate(OUTPUT_COLUMNS, start=1):
        max_len = len(str(col_name))
        for row in rows:
            value = row.get(col_name, "")
            max_len = max(max_len, len(str(value if value is not None else "")))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def main() -> int:
    if len(sys.argv) != 3:
        print("Usage: python extract_expenses.py <root-input-folder> <output-xlsx-path>", file=sys.stderr)
        return 2

    root = Path(sys.argv[1]).resolve()
    output_path = Path(sys.argv[2]).resolve()

    if not root.is_dir():
        print(f"Error: {root} is not a directory", file=sys.stderr)
        return 2

    if CALAMINE_AVAILABLE:
        print("Fallback reader (python-calamine): available")
    else:
        print("Fallback reader (python-calamine): NOT installed — files with")
        print("  malformed XML will be logged as errors. Run:")
        print("  pip install python-calamine")

    error_log_path = output_path.with_name(output_path.stem + "_errors.log")

    files = list(find_excel_files(root))
    print(f"Found {len(files)} Excel file(s) under {root}")
    if not files:
        print("Nothing to do.")
        return 0

    rows: list[dict] = []
    errors: list[str] = []
    method_counts = {"openpyxl": 0, "calamine": 0}

    for i, file_path in enumerate(files, start=1):
        rel = file_path.relative_to(root)
        category, sub_category = derive_categories(rel)
        try:
            fields, method = extract_one(file_path)
            method_counts[method] = method_counts.get(method, 0) + 1
            rows.append({
                "Category": category,
                "Sub-Category": sub_category,
                "Source File": file_path.name,
                "Source Path": str(rel),
                "Read Method": method,
                **fields,
            })
        except Exception as e:
            errors.append(f"{rel}\t{e}")

        if i % 50 == 0 or i == len(files):
            print(f"  Processed {i}/{len(files)}...")

    print(f"Writing {len(rows)} rows to {output_path}")
    print(f"  via openpyxl: {method_counts.get('openpyxl', 0)}")
    print(f"  via calamine fallback: {method_counts.get('calamine', 0)}")
    write_output(rows, output_path)

    if errors:
        with open(error_log_path, "w", encoding="utf-8") as f:
            f.write(f"Errors during extraction ({len(errors)} file(s)):\n\n")
            for line in errors:
                f.write(line + "\n")
        print(f"  {len(errors)} file(s) failed — see {error_log_path}")
    else:
        print("  No errors.")

    print("Done.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
