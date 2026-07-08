"""Config-driven header -> canonical field mapping.

Crosswalk files are small, human-edited JSON documents:

    {
      "canonical_field": {
        "aliases": ["Raw Header Text As Seen In A Real File", "..."],
        "after": "other_canonical_field"   # optional, see below
      }
    }

Matching is exact (after whitespace normalization), never fuzzy: a header
either matches one of the literal aliases on file or it doesn't. Adding
support for a new customer's wording means adding a new alias string to
the JSON, not touching engine code.

Some source formats reuse the exact same header text for several distinct
columns (e.g. a bare "Currency" column that follows several different rate
columns). When an alias string is registered under more than one canonical
field, ``after`` disambiguates: the candidate whose ``after`` equals the
canonical field resolved for the immediately preceding column wins. If no
candidate's ``after`` matches, the column is left unmapped rather than
guessed at.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path

from consolidation.engine.header import normalize_text


@dataclass
class ColumnMapping:
    col_index: int
    raw_header: str
    canonical_field: str | None
    status: str  # "mapped" | "unmapped" | "ambiguous_unresolved" | "blank"


@dataclass
class Crosswalk:
    # normalized alias (lowercased) -> list of (canonical_field, after)
    alias_index: dict[str, list[tuple[str, str | None]]] = field(default_factory=dict)
    canonical_fields: list[str] = field(default_factory=list)

    @classmethod
    def load(cls, path: str | Path) -> "Crosswalk":
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        alias_index: dict[str, list[tuple[str, str | None]]] = {}
        canonical_fields = list(data.keys())
        for canonical_field, spec in data.items():
            after = spec.get("after")
            for alias in spec.get("aliases", []):
                key = normalize_text(alias).lower()
                alias_index.setdefault(key, []).append((canonical_field, after))
        return cls(alias_index=alias_index, canonical_fields=canonical_fields)


def map_headers(raw_headers: list[object], crosswalk: Crosswalk) -> list[ColumnMapping]:
    """Resolve each raw header (in left-to-right order) to a canonical
    field using the crosswalk, disambiguating repeated alias text via the
    canonical field resolved for the previous column."""
    results: list[ColumnMapping] = []
    previous_canonical: str | None = None

    for idx, raw in enumerate(raw_headers, start=1):
        text = normalize_text(raw)
        # A blank header can still be a registered alias (e.g. an unlabeled
        # leading ID column), so it goes through the same lookup as any
        # other header; only truly unregistered blanks get "blank" status.
        candidates = crosswalk.alias_index.get(text.lower(), [])
        if not candidates:
            status = "blank" if not text else "unmapped"
            results.append(
                ColumnMapping(col_index=idx, raw_header=text, canonical_field=None, status=status)
            )
            # a blank header does not reset "previous" — it carries no signal either way
            if text:
                previous_canonical = None
            continue

        if len(candidates) == 1:
            canonical_field, _after = candidates[0]
            results.append(
                ColumnMapping(col_index=idx, raw_header=text, canonical_field=canonical_field, status="mapped")
            )
            previous_canonical = canonical_field
            continue

        # Ambiguous alias: pick the candidate whose `after` matches the
        # previously resolved column's canonical field.
        resolved = None
        for canonical_field, after in candidates:
            if after is not None and after == previous_canonical:
                resolved = canonical_field
                break

        if resolved is not None:
            results.append(
                ColumnMapping(col_index=idx, raw_header=text, canonical_field=resolved, status="mapped")
            )
            previous_canonical = resolved
        else:
            results.append(
                ColumnMapping(col_index=idx, raw_header=text, canonical_field=None, status="ambiguous_unresolved")
            )
            previous_canonical = None

    return results
