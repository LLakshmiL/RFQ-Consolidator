import unittest
from pathlib import Path

import openpyxl

from sync_studio.header_detect import locate_header_row

REPO_ROOT = Path(__file__).resolve().parent.parent
SAMPLE2 = REPO_ROOT / "Sample2.xlsx"


def _sheet_from_rows(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    return ws


class TestLocateHeaderRowSynthetic(unittest.TestCase):
    def test_finds_header_below_metadata_and_description_block(self):
        ws = _sheet_from_rows(
            [
                ["Matrix:", "Some Matrix Name"],
                [None],
                ["Description:"],
                ["This is a very long free-text paragraph explaining the matrix in great detail and should not be mistaken for a header row because it is long."],
                [None],
                ["Service package", "Field description", "Zeilenbeschreibung"],
                ["40'Box/HC", "Offer- / price field", "China"],
                ["20'Box", "Offer- / price field", "China"],
            ]
        )
        match = locate_header_row(ws)
        self.assertEqual(match.row_index, 6)
        self.assertEqual(match.confidence, "high")

    def test_defaults_to_row_one_when_nothing_looks_like_a_header(self):
        ws = _sheet_from_rows([[None, None], [None, None]])
        match = locate_header_row(ws)
        self.assertEqual(match.row_index, 1)

    def test_single_wide_label_row_is_not_mistaken_for_header(self):
        ws = _sheet_from_rows(
            [
                ["Title only, one long cell spanning the whole row conceptually"],
                [None],
                ["Origin", "Destination", "Rate"],
                ["CN", "DE", 100],
            ]
        )
        match = locate_header_row(ws)
        self.assertEqual(match.row_index, 3)


@unittest.skipUnless(SAMPLE2.exists(), "Sample2.xlsx not present")
class TestLocateHeaderRowRealFile(unittest.TestCase):
    def test_matches_known_header_rows_with_no_keyword_hints(self):
        # Ground truth from the earlier keyword/crosswalk-based engine
        # (consolidation/profiles/bid_matrix.py's anchor keywords) — this
        # heuristic gets there with zero knowledge of "Service package" etc.
        expected = {
            "Asia ex Germany": 13,
            "NESEA ex Germany": 13,
            "Middle East and Africa ex Germ_": 15,
            "Middle East and Africa ex China": 15,
            "Northern Europe ex China": 13,
        }
        wb = openpyxl.load_workbook(SAMPLE2, data_only=True)
        for sheet, expected_row in expected.items():
            match = locate_header_row(wb[sheet])
            self.assertEqual(match.row_index, expected_row, msg=f"sheet={sheet}")
            self.assertEqual(match.confidence, "high", msg=f"sheet={sheet}")


if __name__ == "__main__":
    unittest.main()
