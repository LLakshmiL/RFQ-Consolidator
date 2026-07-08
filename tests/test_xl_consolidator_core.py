import io
import unittest
from pathlib import Path

import openpyxl
import pandas as pd

from xl_consolidator.core import (
    ROW_TYPE_COL,
    SOURCE_TAB_COL,
    _col_key,
    _tab_is_source,
    consolidate_workbook,
    export_consolidated_sheet,
    merge_preview,
    norm_text,
    write_back,
)

REPO_ROOT = Path(__file__).resolve().parent.parent
SAMPLE2 = REPO_ROOT / "Sample2.xlsx"

HEADER_ROW = ["Service package", "Field description", "1.1 Pre Carriage", "Quantity", "1.2 DG Surcharge", "Quantity"]


def _make_workbook(sheets: dict[str, list[list]]) -> bytes:
    """Build an in-memory workbook. Each sheet's rows are written starting
    at row 1 (the header row) - `_find_header_row` locates it by the
    literal "Service package" anchor cell, not by position, so a couple of
    blank rows above it (present in real Sample2.xlsx tabs) aren't needed
    for these unit tests."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for row in rows:
            ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _offer_row(*values):
    """A data row whose 'Field description' text contains 'Offer', so it
    survives the ROW_INCLUDE_SUBSTRING filter."""
    return list(values)


class TestNormText(unittest.TestCase):
    def test_strips_leading_number_prefix(self):
        self.assertEqual(norm_text("4.18 E-Mail address"), "e-mail address")
        self.assertEqual(norm_text("4.19 E-Mail address"), "e-mail address")

    def test_collapses_whitespace_and_lowercases(self):
        self.assertEqual(norm_text("  Origin   Country  "), "origin country")

    def test_leaves_text_without_number_dot_pattern_unchanged_besides_case(self):
        self.assertEqual(norm_text("Origin Country"), "origin country")
        self.assertEqual(norm_text("123 Something"), "123 something")  # no dot -> not stripped


class TestColKey(unittest.TestCase):
    def test_same_text_different_number_merges(self):
        self.assertEqual(_col_key("4.18 E-Mail address", ""), _col_key("4.19 E-Mail address", ""))

    def test_same_number_different_text_stays_separate(self):
        self.assertNotEqual(
            _col_key("3.2 Customs Clearance", ""), _col_key("3.2 On Carriage Tokyo port", "")
        )

    def test_quantity_keyed_by_normalized_left_header(self):
        key_a = _col_key("Quantity", "1.2 DG Surcharge Pre Carriage")
        key_b = _col_key("Quantity", "2.2 DG Surcharge Pre Carriage")
        self.assertEqual(key_a, key_b)  # same fee text, different section number -> merges
        self.assertNotEqual(key_a, _col_key("Quantity", "1.3 Multistopp 2nd loading"))


class TestTabIsSource(unittest.TestCase):
    def test_excludes_formulas_prefixed_tabs(self):
        self.assertFalse(_tab_is_source("Formulas_90ca9cbd-6b36-40a0-8c_"))

    def test_excludes_backup_prefixed_tabs(self):
        self.assertFalse(_tab_is_source("_bk_Asia ex Germany"))

    def test_excludes_consolidated_output_sheet(self):
        self.assertFalse(_tab_is_source("Consolidated"))

    def test_includes_ordinary_tab(self):
        self.assertTrue(_tab_is_source("Asia ex Germany"))


class TestConsolidateWorkbook(unittest.TestCase):
    def test_columns_with_differing_section_numbers_merge_into_one(self):
        original_bytes = _make_workbook(
            {
                "A": [
                    ["Service package", "Field description", "4.18 E-Mail address"],
                    ["40'Box", "Offer- / price field", "a@example.com"],
                ],
                "B": [
                    ["Service package", "Field description", "4.19 E-Mail address"],
                    ["20'Box", "Offer- / price field", "b@example.com"],
                ],
            }
        )
        result = consolidate_workbook(original_bytes)
        data_cols = [c for c in result.consolidated.columns if c not in (SOURCE_TAB_COL, ROW_TYPE_COL)]
        self.assertEqual(data_cols, ["Service package", "Field description", "4.18 E-Mail address"])
        values = result.consolidated["4.18 E-Mail address"].tolist()
        self.assertEqual(values, ["a@example.com", "b@example.com"])

    def test_quantity_columns_merge_by_normalized_preceding_header(self):
        original_bytes = _make_workbook(
            {
                "A": [
                    ["Service package", "Field description", "1.2 DG Surcharge", "Quantity"],
                    ["40'Box", "Offer- / price field", 10, 5],
                ],
                "B": [
                    ["Service package", "Field description", "2.2 DG Surcharge", "Quantity"],
                    ["20'Box", "Offer- / price field", 20, 7],
                ],
            }
        )
        result = consolidate_workbook(original_bytes)
        self.assertIn("Quantity [1.2 DG Surcharge]", result.consolidated.columns)
        self.assertNotIn("Quantity [2.2 DG Surcharge]", result.consolidated.columns)
        self.assertEqual(result.consolidated["Quantity [1.2 DG Surcharge]"].tolist(), [5, 7])

    def test_only_offer_rows_are_kept(self):
        original_bytes = _make_workbook(
            {
                "A": [
                    ["Service package", "Field description", "Rate"],
                    ["40'Box", "Offer- / price field", 100],
                    ["40'Box", "Cost- / internal field", 999],
                ]
            }
        )
        result = consolidate_workbook(original_bytes)
        self.assertEqual(len(result.consolidated), 1)
        self.assertEqual(result.consolidated.iloc[0]["Rate"], 100)

    def test_formulas_prefixed_tab_is_excluded(self):
        original_bytes = _make_workbook(
            {
                "Formulas_helper": [["Service package", "Field description"], ["x", "Offer"]],
                "A": [["Service package", "Field description", "Rate"], ["40'Box", "Offer- / price field", 100]],
            }
        )
        result = consolidate_workbook(original_bytes)
        self.assertEqual(set(result.consolidated[SOURCE_TAB_COL]), {"A"})

    def test_provenance_points_to_the_exact_source_cell(self):
        original_bytes = _make_workbook(
            {"A": [["Service package", "Field description", "Rate"], ["40'Box", "Offer- / price field", 100]]}
        )
        result = consolidate_workbook(original_bytes)
        prov = result.provenance[(0, "Rate")]
        self.assertEqual((prov.source_tab, prov.source_row, prov.source_col), ("A", 2, 3))

    def test_raises_if_no_tab_has_the_header_anchor(self):
        original_bytes = _make_workbook({"A": [["Not", "The", "Anchor"], [1, 2, 3]]})
        with self.assertRaises(ValueError):
            consolidate_workbook(original_bytes)


class TestMergePreview(unittest.TestCase):
    def setUp(self):
        self.original_bytes = _make_workbook(
            {"A": [["Service package", "Field description", "Rate"], ["40'Box", "Offer- / price field", 100]]}
        )
        self.result = consolidate_workbook(self.original_bytes)

    def test_no_edits_means_no_writes(self):
        preview = merge_preview(self.result.consolidated.copy(), self.result)
        self.assertFalse(preview.blocked)
        self.assertTrue(preview.writes.empty)
        self.assertTrue(preview.orphans.empty)

    def test_real_edit_is_a_write(self):
        edited = self.result.consolidated.copy()
        edited.loc[0, "Rate"] = 250
        preview = merge_preview(edited, self.result)
        self.assertEqual(len(preview.writes), 1)
        row = preview.writes.iloc[0]
        self.assertEqual((row["SrcTab"], row["OldValue"], row["NewValue"]), ("A", 100, 250))

    def test_blank_edit_does_not_erase_old_value(self):
        edited = self.result.consolidated.copy()
        edited.loc[0, "Rate"] = None
        preview = merge_preview(edited, self.result)
        self.assertTrue(preview.writes.empty)
        self.assertEqual(len(preview._rows("BLANK_MAPPED")), 1)

    def test_pandas_na_is_treated_as_blank(self):
        edited = self.result.consolidated.copy()
        edited["Rate"] = edited["Rate"].astype("string")
        edited.loc[0, "Rate"] = pd.NA
        preview = merge_preview(edited, self.result)
        self.assertTrue(preview.writes.empty)

    def test_reshaping_the_table_blocks_merge(self):
        edited = self.result.consolidated.copy()
        edited["Extra Column"] = "x"
        preview = merge_preview(edited, self.result)
        self.assertTrue(preview.blocked)
        self.assertIn("changed shape", preview.blocked_reason)


class TestMergePreviewFormulas(unittest.TestCase):
    def test_edit_targeting_a_formula_cell_is_skipped_not_written(self):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        ws = wb.create_sheet("A")
        ws.append(["Service package", "Field description", "Rate"])
        ws.append(["40'Box", "Offer- / price field", "=1+1"])
        buf = io.BytesIO()
        wb.save(buf)
        original_bytes = buf.getvalue()

        result = consolidate_workbook(original_bytes)
        edited = result.consolidated.copy()
        edited.loc[0, "Rate"] = 999
        preview = merge_preview(edited, result)
        self.assertTrue(preview.writes.empty)
        self.assertEqual(len(preview.skipped_formulas), 1)


class TestWriteBack(unittest.TestCase):
    def test_writes_value_and_highlights_cell(self):
        original_bytes = _make_workbook(
            {"A": [["Service package", "Field description", "Rate"], ["40'Box", "Offer- / price field", 100]]}
        )
        result = consolidate_workbook(original_bytes)
        edited = result.consolidated.copy()
        edited.loc[0, "Rate"] = 250
        preview = merge_preview(edited, result)

        out_bytes, audit = write_back(original_bytes, result.provenance, preview.writes)

        wb = openpyxl.load_workbook(io.BytesIO(out_bytes))
        cell = wb["A"].cell(row=2, column=3)
        self.assertEqual(cell.value, 250)
        self.assertEqual(cell.fill.fgColor.rgb, "00FFFF00")
        self.assertEqual(audit.iloc[0]["cells_written"], 1)

    def test_original_bytes_are_not_mutated(self):
        original_bytes = _make_workbook(
            {"A": [["Service package", "Field description", "Rate"], ["40'Box", "Offer- / price field", 100]]}
        )
        result = consolidate_workbook(original_bytes)
        edited = result.consolidated.copy()
        edited.loc[0, "Rate"] = 250
        preview = merge_preview(edited, result)
        write_back(original_bytes, result.provenance, preview.writes)

        wb = openpyxl.load_workbook(io.BytesIO(original_bytes))
        self.assertEqual(wb["A"].cell(row=2, column=3).value, 100)


class TestExportConsolidatedSheet(unittest.TestCase):
    def test_adds_consolidated_sheet(self):
        original_bytes = _make_workbook(
            {"A": [["Service package", "Field description", "Rate"], ["40'Box", "Offer- / price field", 100]]}
        )
        result = consolidate_workbook(original_bytes)
        out_bytes = export_consolidated_sheet(original_bytes, result.consolidated)

        wb = openpyxl.load_workbook(io.BytesIO(out_bytes))
        self.assertIn("Consolidated", wb.sheetnames)
        header = [c.value for c in next(wb["Consolidated"].iter_rows(min_row=1, max_row=1))]
        self.assertEqual(header, [SOURCE_TAB_COL, ROW_TYPE_COL, "Service package", "Field description", "Rate"])


@unittest.skipUnless(SAMPLE2.exists(), "Sample2.xlsx not present")
class TestRealWorkbookEndToEnd(unittest.TestCase):
    """Sample2.xlsx is the real, messy production workbook this module was
    built to match: 21 ratecard tabs plus a "Formulas_..." helper tab that
    must be excluded, metadata rows above the real header row, and repeated
    "Quantity" headers that only disambiguate correctly via the preceding
    column's normalized text."""

    def test_consolidate_merge_write_back_round_trip(self):
        original_bytes = SAMPLE2.read_bytes()
        result = consolidate_workbook(original_bytes)

        # The helper "Formulas_..." tab must never appear as a source.
        self.assertNotIn("Formulas_90ca9cbd-6b36-40a0-8c_", set(result.consolidated[SOURCE_TAB_COL]))
        self.assertGreater(len(result.consolidated), 0)

        edited = result.consolidated.copy()
        mask = (edited[SOURCE_TAB_COL] == "Asia ex Germany") & (
            edited["Zeilenbeschreibung / Spaltenbeschreibung"] == "China"
        )
        target_row = edited.index[mask][0]  # first matching row only - the mask matches 2 (one per service package)
        target_col = "Quantity [1.1 Pre Carriage ex Postcode 74xxx (incl. OTHC)]"
        self.assertEqual(edited.loc[target_row, target_col], 173)
        edited.loc[target_row, target_col] = 999999

        preview = merge_preview(edited, result)
        self.assertFalse(preview.blocked)
        self.assertEqual(len(preview.writes), 1)

        out_bytes, _audit = write_back(original_bytes, result.provenance, preview.writes)
        wb = openpyxl.load_workbook(io.BytesIO(out_bytes))
        self.assertEqual(wb["Asia ex Germany"].cell(row=14, column=6).value, 999999)


if __name__ == "__main__":
    unittest.main()
