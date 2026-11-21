import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import extract_expenses as ex          # noqa: E402
import make_sample_claims as sample    # noqa: E402
from openpyxl import load_workbook     # noqa: E402


def build_samples(tmp: Path) -> Path:
    root = tmp / "claims"
    for rec in sample.CLAIMS:
        sample.write_claim(root / rec[0] / rec[1], rec)
    sample.write_claim(root / "Group B/Multan" / "Claim_90099_WrongSheet.xlsx", sample.CLAIMS[0], sheet="Sheet1")
    (root / "Group A/Lahore" / "~$lock.xlsx").write_bytes(b"lock")
    return root


class Helpers(unittest.TestCase):
    def test_address_conversion(self):
        self.assertEqual(ex._addr_to_rc("A1"), (0, 0))
        self.assertEqual(ex._addr_to_rc("B7"), (6, 1))
        self.assertEqual(ex._addr_to_rc("AA10"), (9, 26))

    def test_normalize_matches_openpyxl_conventions(self):
        self.assertIsNone(ex._normalize(""))
        self.assertIsNone(ex._normalize(None))
        self.assertEqual(ex._normalize(12.0), 12)
        self.assertIsInstance(ex._normalize(12.0), int)
        self.assertEqual(ex._normalize(12.5), 12.5)
        self.assertEqual(ex._normalize("text"), "text")

    def test_categories_come_from_the_folder_names(self):
        self.assertEqual(ex.derive_categories(Path("a.xlsx")), ("", ""))
        self.assertEqual(ex.derive_categories(Path("Group A/a.xlsx")), ("Group A", ""))
        self.assertEqual(ex.derive_categories(Path("Group A/Lahore/x/a.xlsx")), ("Group A", "Lahore/x"))


class Reading(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = build_samples(Path(self.tmp.name))
        self.file = self.root / "Group A/Lahore/Claim_90001_Jun-2026.xlsx"

    def tearDown(self):
        self.tmp.cleanup()

    def test_walker_skips_lock_files_hidden_files_and_other_types(self):
        (self.root / ".hidden.xlsx").write_bytes(b"x")
        (self.root / "notes.txt").write_text("x")
        names = sorted(p.name for p in ex.find_excel_files(self.root))
        self.assertEqual(len(names), 12)
        self.assertFalse(any(n.startswith(("~$", ".")) for n in names))
        self.assertTrue(all(n.endswith(".xlsx") for n in names))

    def test_openpyxl_reads_every_mapped_field(self):
        got = ex.read_with_openpyxl(self.file)
        self.assertEqual(set(got), set(ex.FIELD_MAP))
        self.assertEqual(got["Name"], "Imran Sample")
        self.assertEqual(got["Employee ID"], 90001)
        self.assertEqual(got["Total KM"], 1480)
        self.assertEqual(got["Grand Total"], got["Trav. Exp. Total"] + got["Sundries Total"])

    @unittest.skipUnless(ex.CALAMINE_AVAILABLE, "python-calamine is not installed")
    def test_calamine_gives_the_same_values_as_openpyxl(self):
        self.assertEqual(ex.read_with_calamine(self.file), ex.read_with_openpyxl(self.file))

    @unittest.skipUnless(ex.CALAMINE_AVAILABLE, "python-calamine is not installed")
    def test_fallback_reader_is_used_when_openpyxl_hits_malformed_xml(self):
        with mock.patch.object(ex, "read_with_openpyxl", side_effect=ValueError("malformed XML")):
            fields, method = ex.extract_one(self.file)
        self.assertEqual(method, "calamine")
        self.assertEqual(fields["Name"], "Imran Sample")

    def test_a_missing_sheet_is_reported_not_retried(self):
        bad = self.root / "Group B/Multan/Claim_90099_WrongSheet.xlsx"
        with self.assertRaisesRegex(ValueError, "not found"):
            ex.extract_one(bad)


class EndToEnd(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = build_samples(Path(self.tmp.name))
        self.out = Path(self.tmp.name) / "out" / "Consolidated.xlsx"

    def tearDown(self):
        self.tmp.cleanup()

    def run_main(self, *args):
        with mock.patch.object(sys, "argv", ["extract_expenses.py", *map(str, args)]):
            return ex.main()

    def test_one_row_per_file_and_an_error_log_for_the_bad_one(self):
        self.assertEqual(self.run_main(self.root, self.out), 0)
        ws = load_workbook(self.out).active
        header = [c.value for c in ws[1]]
        self.assertEqual(header, ex.OUTPUT_COLUMNS)
        self.assertEqual(ws.max_row - 1, 11)                         # 12 workbooks, 1 with the wrong sheet name
        rows = [dict(zip(header, r)) for r in ws.iter_rows(min_row=2, values_only=True)]
        by_id = {r["Employee ID"]: r for r in rows}
        self.assertEqual(by_id[90006]["Category"], "Group B")
        self.assertEqual(by_id[90006]["Sub-Category"], "Islamabad")
        self.assertEqual(by_id[90011]["Category"] or "", "")         # a file in the root has no category
        self.assertEqual({r["Read Method"] for r in rows}, {"openpyxl"})
        log = (self.out.parent / "Consolidated_errors.log").read_text(encoding="utf-8")
        self.assertIn("Claim_90099_WrongSheet.xlsx", log)
        self.assertIn("sheet 'Traveling Expense' not found", log)

    def test_source_files_are_never_modified(self):
        before = {p: p.stat().st_mtime_ns for p in ex.find_excel_files(self.root)}
        self.run_main(self.root, self.out)
        after = {p: p.stat().st_mtime_ns for p in ex.find_excel_files(self.root)}
        self.assertEqual(before, after)

    def test_usage_errors_and_empty_folder(self):
        self.assertEqual(self.run_main(self.root), 2)                       # wrong number of arguments
        self.assertEqual(self.run_main(self.root / "nope", self.out), 2)    # not a directory
        empty = Path(self.tmp.name) / "empty"
        empty.mkdir()
        self.assertEqual(self.run_main(empty, self.out), 0)                 # nothing to do


if __name__ == "__main__":
    unittest.main()
