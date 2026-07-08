"""A Profile is the unit of reusability: it declares everything the engine
needs to recognize and process one customer file *family* (a group of
files/sheets that share the same layout). Adding support for a new
customer means either adding aliases to an existing profile's crosswalk
JSON (same layout, different wording) or adding a new Profile module (a
genuinely different layout) — the orchestration code in
engine/consolidate.py never needs to change.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from consolidation.engine.mapping import Crosswalk


@dataclass
class Profile:
    name: str
    anchor_keywords: list[str]
    crosswalk_path: Path
    template_columns: list[str]
    kind: str  # "flat" (one output row per data row) | "unpivot" (melt repeating column pairs)
    min_hits: int = 3
    value_header_match: str = "quantity"  # only used when kind == "unpivot"
    extract_sheet_metadata: Callable | None = None  # ws -> dict, optional sheet-level metadata
    text_columns: tuple[str, ...] = ()  # template columns that hold codes, not numbers (e.g. "1.10" != 1.1)

    def load_crosswalk(self) -> Crosswalk:
        return Crosswalk.load(self.crosswalk_path)


_REGISTRY: dict[str, Profile] = {}


def register(profile: Profile) -> Profile:
    _REGISTRY[profile.name] = profile
    return profile


def all_profiles() -> list[Profile]:
    return list(_REGISTRY.values())
