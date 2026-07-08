import unittest

from sync_studio.diff_sync import build_sync_plan
from sync_studio.parser import TabData


def _tab(sheet_name, columns, rows, first_row_number=2):
    row_numbers = list(range(first_row_number, first_row_number + len(rows)))
    return TabData(
        sheet_name=sheet_name,
        header_row=first_row_number - 1,
        detection=None,
        columns=columns,
        rows=rows,
        row_numbers=row_numbers,
    )


class TestBuildSyncPlan(unittest.TestCase):
    def test_modified_cell_detected(self):
        tab = _tab("A", ["Origin", "Rate"], [["CN", 100]])
        edited_rows = [{"__source_tab": "A", "__source_row": 2, "Origin": "CN", "Rate": 150}]
        plan = build_sync_plan([tab], ["Origin", "Rate"], edited_rows)
        tab_plan = plan.tabs[0]
        self.assertEqual(len(tab_plan.modified_cells), 1)
        self.assertEqual(tab_plan.modified_cells[0].row_number, 2)
        self.assertEqual(tab_plan.modified_cells[0].col_index, 2)
        self.assertEqual(tab_plan.modified_cells[0].value, 150)
        self.assertEqual(tab_plan.new_columns, [])
        self.assertEqual(tab_plan.deleted_columns, [])

    def test_unchanged_cell_produces_no_modification(self):
        tab = _tab("A", ["Origin", "Rate"], [["CN", 100]])
        edited_rows = [{"__source_tab": "A", "__source_row": 2, "Origin": "CN", "Rate": 100}]
        plan = build_sync_plan([tab], ["Origin", "Rate"], edited_rows)
        self.assertEqual(plan.tabs[0].modified_cells, [])

    def test_new_column_only_written_for_tab_with_a_value(self):
        tab_a = _tab("A", ["Origin"], [["CN"]])
        tab_b = _tab("B", ["Origin"], [["DE"]])
        edited_rows = [
            {"__source_tab": "A", "__source_row": 2, "Origin": "CN", "Notes": "checked"},
            {"__source_tab": "B", "__source_row": 2, "Origin": "DE", "Notes": ""},
        ]
        plan = build_sync_plan([tab_a, tab_b], ["Origin", "Notes"], edited_rows)
        plan_a = next(t for t in plan.tabs if t.sheet_name == "A")
        plan_b = next(t for t in plan.tabs if t.sheet_name == "B")
        self.assertEqual(len(plan_a.new_columns), 1)
        self.assertEqual(plan_a.new_columns[0].column_name, "Notes")
        self.assertEqual(plan_a.new_columns[0].values, {2: "checked"})
        # tab B has no non-blank value for "Notes" -> no new column written there
        self.assertEqual(plan_b.new_columns, [])

    def test_column_removed_from_master_is_flagged_for_deletion_in_every_tab_that_had_it(self):
        tab_a = _tab("A", ["Origin", "Legacy"], [["CN", "x"]])
        tab_b = _tab("B", ["Origin", "Legacy"], [["DE", "y"]])
        edited_rows = [
            {"__source_tab": "A", "__source_row": 2, "Origin": "CN"},
            {"__source_tab": "B", "__source_row": 2, "Origin": "DE"},
        ]
        # "Legacy" column deleted client-side -> not present in edited_columns at all
        plan = build_sync_plan([tab_a, tab_b], ["Origin"], edited_rows)
        for tab_plan in plan.tabs:
            self.assertEqual(tab_plan.deleted_columns, ["Legacy"])

    def test_type_only_roundtrip_noise_is_not_flagged_as_modified(self):
        # A dataframe-backed grid can turn an original int cell into a
        # string for a row nobody touched (see diff_sync._as_float's
        # docstring) — must not be reported as a user edit.
        tab = _tab("A", ["Bulkiness"], [[1]])
        edited_rows = [{"__source_tab": "A", "__source_row": 2, "Bulkiness": "1"}]
        plan = build_sync_plan([tab], ["Bulkiness"], edited_rows)
        self.assertEqual(plan.tabs[0].modified_cells, [])

    def test_genuine_numeric_change_is_still_flagged(self):
        tab = _tab("A", ["Rate"], [[100]])
        edited_rows = [{"__source_tab": "A", "__source_row": 2, "Rate": "150"}]
        plan = build_sync_plan([tab], ["Rate"], edited_rows)
        self.assertEqual(len(plan.tabs[0].modified_cells), 1)
        self.assertEqual(plan.tabs[0].modified_cells[0].value, "150")

    def test_duplicate_same_tab_header_columns_diffed_independently(self):
        tab = _tab("A", ["Fee", "Currency", "Currency"], [[10, "USD", "EUR"]])
        edited_rows = [{"__source_tab": "A", "__source_row": 2, "Fee": 10, "Currency": "GBP", "Currency (2)": "EUR"}]
        plan = build_sync_plan([tab], ["Fee", "Currency", "Currency (2)"], edited_rows)
        modified = {c.col_index: c.value for c in plan.tabs[0].modified_cells}
        self.assertEqual(modified, {2: "GBP"})  # only the first "Currency" (col 2) changed


if __name__ == "__main__":
    unittest.main()
