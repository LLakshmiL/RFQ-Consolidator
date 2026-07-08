import unittest

from consolidation.engine.mapping import ColumnMapping
from consolidation.engine.unpivot import (
    build_unpivot_plan,
    melt_row,
    split_attribute_code,
)


def _mapping(col_index, raw_header, canonical_field=None, status=None):
    if status is None:
        status = "mapped" if canonical_field else "unmapped"
    return ColumnMapping(col_index=col_index, raw_header=raw_header, canonical_field=canonical_field, status=status)


class TestSplitAttributeCode(unittest.TestCase):
    def test_splits_numeric_prefix(self):
        self.assertEqual(
            split_attribute_code("1.1 Pre Carriage ex Postcode 74xxx (incl. OTHC)"),
            ("1.1", "Pre Carriage ex Postcode 74xxx (incl. OTHC)"),
        )

    def test_no_prefix_keeps_full_text(self):
        self.assertEqual(split_attribute_code("Quantity"), ("", "Quantity"))


class TestBuildUnpivotPlan(unittest.TestCase):
    def test_variable_length_pairs_after_fixed_identifier_columns(self):
        raw_headers = [
            None,
            "Service package",
            "Field description",
            "1.1 Pre Carriage",
            "Quantity",
            "1.2 DG Surcharge",
            "Quantity",
        ]
        column_mappings = [
            _mapping(1, "", "record_id"),
            _mapping(2, "Service package", "service_package"),
            _mapping(3, "Field description", "field_description"),
            _mapping(4, "1.1 Pre Carriage"),
            _mapping(5, "Quantity"),
            _mapping(6, "1.2 DG Surcharge"),
            _mapping(7, "Quantity"),
        ]
        plan = build_unpivot_plan(raw_headers, column_mappings)
        self.assertEqual(plan.identifier_end_col, 3)
        self.assertEqual(len(plan.pairs), 2)
        self.assertEqual(plan.pairs[0].label_col, 4)
        self.assertEqual(plan.pairs[0].value_col, 5)
        self.assertEqual(plan.pairs[1].label_col, 6)
        self.assertEqual(plan.pairs[1].value_col, 7)
        self.assertEqual(plan.leftover_cols, [])

    def test_unpaired_trailing_column_reported_as_leftover_not_dropped(self):
        raw_headers = ["Service package", "1.1 Pre Carriage", "Quantity", "Orphan Column"]
        column_mappings = [
            _mapping(1, "Service package", "service_package"),
            _mapping(2, "1.1 Pre Carriage"),
            _mapping(3, "Quantity"),
            _mapping(4, "Orphan Column"),
        ]
        plan = build_unpivot_plan(raw_headers, column_mappings)
        self.assertEqual(len(plan.pairs), 1)
        self.assertEqual(plan.leftover_cols, [4])

    def test_works_regardless_of_pair_count(self):
        # A sheet with many more repeating pairs than another must still work —
        # this is what makes the melt generic across differently-sized sheets.
        raw_headers = ["Service package"]
        column_mappings = [_mapping(1, "Service package", "service_package")]
        for i in range(1, 21):
            raw_headers += [f"{i}.1 Charge {i}", "Quantity"]
            column_mappings += [_mapping(len(raw_headers) - 1, raw_headers[-2]), _mapping(len(raw_headers), "Quantity")]
        plan = build_unpivot_plan(raw_headers, column_mappings)
        self.assertEqual(len(plan.pairs), 20)
        self.assertEqual(plan.leftover_cols, [])


class TestMeltRow(unittest.TestCase):
    def test_emits_one_row_per_non_blank_pair_and_skips_blank_values(self):
        raw_headers = ["Service package", "1.1 Pre Carriage", "Quantity", "1.2 DG Surcharge", "Quantity"]
        column_mappings = [
            _mapping(1, "Service package", "service_package"),
            _mapping(2, "1.1 Pre Carriage"),
            _mapping(3, "Quantity"),
            _mapping(4, "1.2 DG Surcharge"),
            _mapping(5, "Quantity"),
        ]
        plan = build_unpivot_plan(raw_headers, column_mappings)
        # positions: [service_package, label1(blank placeholder), value1, label2(blank placeholder), value2]
        row_values = ["40'", "", 173, "", 0]
        melted = melt_row(row_values, plan)
        self.assertEqual(len(melted), 2)
        self.assertEqual(melted[0]["attribute_code"], "1.1")
        self.assertEqual(melted[0]["attribute_value"], 173)
        self.assertEqual(melted[1]["attribute_code"], "1.2")
        self.assertEqual(melted[1]["attribute_value"], 0)

    def test_blank_value_pair_is_skipped(self):
        raw_headers = ["1.1 Pre Carriage", "Quantity"]
        column_mappings = [_mapping(1, "1.1 Pre Carriage"), _mapping(2, "Quantity")]
        plan = build_unpivot_plan(raw_headers, column_mappings)
        melted = melt_row(["", None], plan)
        self.assertEqual(melted, [])


if __name__ == "__main__":
    unittest.main()
