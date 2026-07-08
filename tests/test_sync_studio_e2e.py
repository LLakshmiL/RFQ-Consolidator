import io
import unittest
from pathlib import Path

import openpyxl

from sync_studio.diff_sync import build_sync_plan
from sync_studio.header_detect import normalize_text
from sync_studio.master import build_master
from sync_studio.parser import parse_workbook
from sync_studio.writer import apply_sync_plan

REPO_ROOT = Path(__file__).resolve().parent.parent
SAMPLE1 = REPO_ROOT / "Sample1.xlsx"


@unittest.skipUnless(SAMPLE1.exists(), "Sample1.xlsx not present")
class TestSyncStudioEndToEnd(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.original_bytes = SAMPLE1.read_bytes()
        wb = openpyxl.load_workbook(io.BytesIO(cls.original_bytes), data_only=True)
        cls.baseline_tabs = parse_workbook(wb)
        cls.master = build_master(cls.baseline_tabs)

    def _edited_rows_with(self, mutate):
        rows = [dict(row) for row in self.master.rows]
        mutate(rows)
        return rows

    def test_modify_add_and_delete_column_round_trip(self):
        target_tab = "AP - EU LCL Ratecard"
        target_row = next(r for r in self.master.rows if r["__source_tab"] == target_tab)
        target_row_number = target_row["__source_row"]

        columns = list(self.master.columns)
        # 1. modify an existing cell
        modify_col = "10 No. Of Cbm [1 Year Volume]" if "10 No. Of Cbm [1 Year Volume]" in columns else columns[9]
        # find the real volume column name actually present
        volume_col = next(c for c in columns if "No. Of Cbm" in c)

        # 2. add a new column, give it a value only for our target row
        columns.append("Reviewed By")

        # 3. delete a column that exists in every Sample1 tab
        delete_col = "36.1 Total Transit Time [days]"
        self.assertIn(delete_col, columns)
        columns.remove(delete_col)

        def mutate(rows):
            for row in rows:
                row.pop(delete_col, None)
                row["Reviewed By"] = ""
            for row in rows:
                if row["__source_tab"] == target_tab and row["__source_row"] == target_row_number:
                    row[volume_col] = 999
                    row["Reviewed By"] = "Alex"

        edited_rows = self._edited_rows_with(mutate)

        plan = build_sync_plan(self.baseline_tabs, columns, edited_rows)
        result_bytes = apply_sync_plan(self.original_bytes, self.baseline_tabs, plan)

        out_wb = openpyxl.load_workbook(io.BytesIO(result_bytes))

        # --- modified cell: correct value + yellow highlight ---
        ws = out_wb[target_tab]
        # The original header cell's raw text (e.g. a trailing space) is left
        # untouched by writer.py since it isn't a cell we're changing — so
        # this lookup normalizes before matching, same as parser.py does.
        header = [normalize_text(c.value) for c in ws[6]]
        volume_idx = header.index(volume_col) + 1
        modified_cell = ws.cell(row=target_row_number, column=volume_idx)
        self.assertEqual(modified_cell.value, 999)
        self.assertEqual(modified_cell.fill.start_color.rgb, "00FFFF00")

        # --- new column: appended only to the target tab, highlighted ---
        new_header = [c.value for c in ws[6]]
        self.assertIn("Reviewed By", new_header)
        new_col_idx = new_header.index("Reviewed By") + 1
        new_cell = ws.cell(row=target_row_number, column=new_col_idx)
        self.assertEqual(new_cell.value, "Alex")
        self.assertEqual(new_cell.fill.start_color.rgb, "00FFFF00")

        other_tab_ws = out_wb["EU - NA LCL Ratecard"]
        other_header = [c.value for c in other_tab_ws[6]]
        self.assertNotIn("Reviewed By", other_header)

        # --- deleted column: gone from every tab that had it ---
        for sheet_name in ["AP - EU LCL Ratecard", "EU - NA LCL Ratecard", "AP - NA LCL Ratecard"]:
            header_texts = [c.value for c in out_wb[sheet_name][6]]
            self.assertNotIn(delete_col, header_texts)

        # --- Audit sheet exists and mentions the changes ---
        self.assertIn("Audit", out_wb.sheetnames)
        audit_ws = out_wb["Audit"]
        audit_rows = list(audit_ws.iter_rows(values_only=True))
        self.assertEqual(audit_rows[0], ("Sheet", "Cells Modified", "Columns Added", "Columns Removed"))
        target_audit_row = next(r for r in audit_rows[1:] if r[0] == target_tab)
        self.assertGreaterEqual(target_audit_row[1], 1)
        self.assertIn("Reviewed By", target_audit_row[2])
        self.assertIn(delete_col, target_audit_row[3])


if __name__ == "__main__":
    unittest.main()
