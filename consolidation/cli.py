"""CLI entry point.

    python -m consolidation.cli Sample1.xlsx Sample2.xlsx --out output/

Accepts any number of workbook paths. Unrecognized files/sheets are
reported in the audit log, not treated as fatal errors.
"""
from __future__ import annotations

import argparse
import sys

from consolidation.engine.consolidate import consolidate


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Consolidate multi-worksheet Excel workbooks.")
    parser.add_argument("inputs", nargs="+", help="Path(s) to input .xlsx workbook(s)")
    parser.add_argument("--out", default="output", help="Output directory (default: output/)")
    args = parser.parse_args(argv)

    audit_log = consolidate(args.inputs, args.out)

    total = len(audit_log.entries)
    processed = sum(1 for e in audit_log.entries if e.status == "processed")
    skipped = total - processed
    print(f"Processed {processed}/{total} sheet(s); {skipped} skipped (see audit_report.xlsx).")
    for e in audit_log.entries:
        if e.status != "processed":
            print(f"  SKIPPED  {e.source_file} :: {e.source_sheet}  ({e.status})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
