import unittest

import openpyxl

from consolidation.engine.header import locate_header_row, normalize_text, score_sheet


def _sheet_from_rows(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    return ws


class TestNormalizeText(unittest.TestCase):
    def test_collapses_whitespace_and_trims(self):
        self.assertEqual(normalize_text("  26  Import   Customs "), "26 Import Customs")

    def test_none_becomes_empty_string(self):
        self.assertEqual(normalize_text(None), "")


class TestLocateHeaderRow(unittest.TestCase):
    def test_finds_header_below_metadata_block(self):
        ws = _sheet_from_rows(
            [
                ["Title", None, None],
                ["Rates valid from", "2026-01-01", None],
                [None, None, None],
                ["Origin Country", "Port of Loading", "Destination Country"],
                ["CN", "Shanghai", "DE"],
            ]
        )
        match = locate_header_row(
            ws, ["Origin Country", "Port of Loading", "Destination Country"], min_hits=2
        )
        self.assertIsNotNone(match)
        self.assertEqual(match.row_index, 4)
        self.assertEqual(match.hits, 3)

    def test_returns_none_when_no_row_reaches_min_hits(self):
        ws = _sheet_from_rows([["Unrelated", "Columns", "Here"]])
        match = locate_header_row(ws, ["Origin Country", "Port of Loading"], min_hits=2)
        self.assertIsNone(match)

    def test_header_row_position_varies_and_is_still_found(self):
        # Same anchors, but pushed two rows further down by extra metadata —
        # this is exactly the Sample2 "Middle East" sheets' row-13-vs-15 case.
        ws = _sheet_from_rows(
            [
                ["Matrix:", "Name"],
                [None],
                [None],
                ["Service package", "Field description"],
            ]
        )
        match = locate_header_row(ws, ["Service package", "Field description"], min_hits=2)
        self.assertEqual(match.row_index, 4)


class TestScoreSheet(unittest.TestCase):
    def test_score_reflects_best_anchor_hit_count(self):
        ws = _sheet_from_rows([["Origin Country", "Port of Loading", "Unrelated"]])
        score = score_sheet(ws, ["Origin Country", "Port of Loading", "Destination Country"])
        self.assertEqual(score, 2)


if __name__ == "__main__":
    unittest.main()
