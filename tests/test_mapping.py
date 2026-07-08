import unittest

from consolidation.engine.mapping import Crosswalk, map_headers


def _crosswalk(spec: dict) -> Crosswalk:
    """Build a Crosswalk in-memory, same shape as Crosswalk.load() would
    produce from a JSON file, without touching the filesystem."""
    alias_index: dict[str, list[tuple[str, str | None]]] = {}
    for canonical_field, entry in spec.items():
        after = entry.get("after")
        for alias in entry["aliases"]:
            alias_index.setdefault(alias.lower(), []).append((canonical_field, after))
    return Crosswalk(alias_index=alias_index, canonical_fields=list(spec.keys()))


class TestMapHeaders(unittest.TestCase):
    def test_exact_alias_match(self):
        xwalk = _crosswalk({"origin_country": {"aliases": ["1 Origin Country [ISO country code]"]}})
        result = map_headers(["1 Origin Country [ISO country code]"], xwalk)
        self.assertEqual(result[0].canonical_field, "origin_country")
        self.assertEqual(result[0].status, "mapped")

    def test_whitespace_variants_still_match(self):
        xwalk = _crosswalk({"fee": {"aliases": ["26 Import Customs Clearance"]}})
        result = map_headers(["  26   Import  Customs   Clearance  "], xwalk)
        self.assertEqual(result[0].canonical_field, "fee")

    def test_unregistered_header_is_unmapped_not_dropped(self):
        xwalk = _crosswalk({"origin_country": {"aliases": ["Origin Country"]}})
        result = map_headers(["Some New Column Nobody Mapped Yet"], xwalk)
        self.assertIsNone(result[0].canonical_field)
        self.assertEqual(result[0].status, "unmapped")
        self.assertEqual(result[0].raw_header, "Some New Column Nobody Mapped Yet")

    def test_ambiguous_alias_disambiguated_by_after(self):
        xwalk = _crosswalk(
            {
                "consolidation_fee": {"aliases": ["Consolidation Fee"]},
                "consolidation_fee_currency": {"aliases": ["Currency"], "after": "consolidation_fee"},
                "deconsolidation_fee": {"aliases": ["Deconsolidation Fee"]},
                "deconsolidation_fee_currency": {"aliases": ["Currency"], "after": "deconsolidation_fee"},
            }
        )
        result = map_headers(
            ["Consolidation Fee", "Currency", "Deconsolidation Fee", "Currency"], xwalk
        )
        self.assertEqual(result[1].canonical_field, "consolidation_fee_currency")
        self.assertEqual(result[1].status, "mapped")
        self.assertEqual(result[3].canonical_field, "deconsolidation_fee_currency")
        self.assertEqual(result[3].status, "mapped")

    def test_ambiguous_alias_without_matching_predecessor_is_unresolved(self):
        xwalk = _crosswalk(
            {
                "consolidation_fee": {"aliases": ["Consolidation Fee"]},
                "consolidation_fee_currency": {"aliases": ["Currency"], "after": "consolidation_fee"},
                "deconsolidation_fee": {"aliases": ["Deconsolidation Fee"]},
                "deconsolidation_fee_currency": {"aliases": ["Currency"], "after": "deconsolidation_fee"},
            }
        )
        # "Currency" with no recognized field before it — must not guess.
        result = map_headers(["Currency"], xwalk)
        self.assertIsNone(result[0].canonical_field)
        self.assertEqual(result[0].status, "ambiguous_unresolved")

    def test_blank_header_matches_registered_blank_alias(self):
        xwalk = _crosswalk({"record_id": {"aliases": [""]}})
        result = map_headers([None], xwalk)
        self.assertEqual(result[0].canonical_field, "record_id")
        self.assertEqual(result[0].status, "mapped")

    def test_unregistered_blank_header_is_status_blank(self):
        xwalk = _crosswalk({"origin_country": {"aliases": ["Origin Country"]}})
        result = map_headers([None], xwalk)
        self.assertIsNone(result[0].canonical_field)
        self.assertEqual(result[0].status, "blank")


if __name__ == "__main__":
    unittest.main()
