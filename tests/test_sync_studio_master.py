import unittest

from sync_studio.master import build_master, effective_columns
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


class TestEffectiveColumns(unittest.TestCase):
    def test_duplicate_headers_in_same_tab_are_disambiguated(self):
        result = effective_columns(["Name", "Currency", "Currency", "Currency"])
        self.assertEqual(result, ["Name", "Currency", "Currency (2)", "Currency (3)"])

    def test_blank_headers_pass_through_unchanged(self):
        self.assertEqual(effective_columns(["", "Name", ""]), ["", "Name", ""])


class TestBuildMaster(unittest.TestCase):
    def test_outer_union_of_columns_across_tabs(self):
        tab_a = _tab("A", ["Origin", "Rate"], [["CN", 100]])
        tab_b = _tab("B", ["Origin", "Destination"], [["DE", "US"]])
        master = build_master([tab_a, tab_b])
        self.assertEqual(master.columns, ["Origin", "Rate", "Destination"])
        self.assertEqual(len(master.rows), 2)
        row_a = next(r for r in master.rows if r["__source_tab"] == "A")
        row_b = next(r for r in master.rows if r["__source_tab"] == "B")
        self.assertEqual(row_a, {"__source_tab": "A", "__source_row": 2, "Origin": "CN", "Rate": 100, "Destination": None})
        self.assertEqual(row_b, {"__source_tab": "B", "__source_row": 2, "Origin": "DE", "Rate": None, "Destination": "US"})

    def test_header_matching_is_case_and_whitespace_insensitive(self):
        tab_a = _tab("A", ["Origin Country"], [["CN"]])
        tab_b = _tab("B", ["  origin   country  "], [["DE"]])
        master = build_master([tab_a, tab_b])
        self.assertEqual(master.columns, ["Origin Country"])
        self.assertEqual([r["Origin Country"] for r in master.rows], ["CN", "DE"])

    def test_same_tab_duplicate_header_columns_both_survive_in_master(self):
        # Sample1.xlsx's real-world case: one tab has 11 columns literally
        # named "Currency" for different fees. Neither may be silently lost.
        tab = _tab("A", ["Fee", "Currency", "OtherFee", "Currency"], [[10, "USD", 20, "EUR"]])
        master = build_master([tab])
        self.assertEqual(master.columns, ["Fee", "Currency", "OtherFee", "Currency (2)"])
        self.assertEqual(master.rows[0]["Currency"], "USD")
        self.assertEqual(master.rows[0]["Currency (2)"], "EUR")

    def test_blank_header_columns_are_excluded_from_master(self):
        tab = _tab("A", ["", "Name"], [[999, "x"]])
        master = build_master([tab])
        self.assertEqual(master.columns, ["Name"])


if __name__ == "__main__":
    unittest.main()
