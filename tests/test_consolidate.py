import tempfile
import unittest
from pathlib import Path

from consolidation.engine.consolidate import consolidate

REPO_ROOT = Path(__file__).resolve().parent.parent
SAMPLE1 = REPO_ROOT / "Sample1.xlsx"
SAMPLE2 = REPO_ROOT / "Sample2.xlsx"


@unittest.skipUnless(SAMPLE1.exists() and SAMPLE2.exists(), "sample workbooks not present")
class TestConsolidateEndToEnd(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmpdir = tempfile.TemporaryDirectory()
        cls.audit_log = consolidate([str(SAMPLE1), str(SAMPLE2)], cls.tmpdir.name)
        cls.entries_by_sheet = {e.source_sheet: e for e in cls.audit_log.entries}

    @classmethod
    def tearDownClass(cls):
        cls.tmpdir.cleanup()

    def test_hidden_helper_sheet_is_skipped_not_guessed_at(self):
        entry = self.entries_by_sheet["Formulas_90ca9cbd-6b36-40a0-8c_"]
        self.assertEqual(entry.status, "skipped_hidden")

    def test_sample1_sheets_detected_as_lcl_ratecard(self):
        for sheet in ["AP - EU LCL Ratecard", "EU - NA LCL Ratecard", "AP - NA LCL Ratecard"]:
            entry = self.entries_by_sheet[sheet]
            self.assertEqual(entry.status, "processed")
            self.assertEqual(entry.profile, "lcl_ratecard")
            self.assertEqual(entry.header_row, 6)
            self.assertEqual(entry.mapped_columns, 38)
            self.assertEqual(entry.unmapped_headers, [])

    def test_sample2_sheets_detected_as_bid_matrix_with_varying_header_row(self):
        # Most sheets have their header at row 13; the "Middle East" sheets
        # have two extra metadata rows, pushing it to row 15. Anchor-based
        # detection must find both without being told where to look.
        self.assertEqual(self.entries_by_sheet["Asia ex Germany"].header_row, 13)
        self.assertEqual(self.entries_by_sheet["Middle East and Africa ex Germ_"].header_row, 15)
        for entry in self.entries_by_sheet.values():
            if entry.profile == "bid_matrix":
                self.assertEqual(entry.status, "processed")
                self.assertEqual(entry.mapped_columns, 4)
                self.assertEqual(entry.unmapped_headers, [])

    def test_no_sheet_left_unrecognized(self):
        for sheet, entry in self.entries_by_sheet.items():
            self.assertIn(
                entry.status,
                ("processed", "skipped_hidden"),
                msg=f"{sheet} was {entry.status}, expected a known profile or a hidden-sheet skip",
            )

    def test_master_outputs_written_with_matching_row_counts(self):
        import pandas as pd

        lcl_df = pd.read_excel(Path(self.tmpdir.name) / "consolidated_lcl_ratecard.xlsx")
        bid_df = pd.read_excel(Path(self.tmpdir.name) / "consolidated_bid_matrix.xlsx")

        expected_lcl_rows = sum(
            e.rows_out for e in self.audit_log.entries if e.profile == "lcl_ratecard"
        )
        expected_bid_rows = sum(
            e.rows_out for e in self.audit_log.entries if e.profile == "bid_matrix"
        )
        self.assertEqual(len(lcl_df), expected_lcl_rows)
        self.assertEqual(len(bid_df), expected_bid_rows)
        self.assertGreater(len(lcl_df), 0)
        self.assertGreater(len(bid_df), 0)


if __name__ == "__main__":
    unittest.main()
