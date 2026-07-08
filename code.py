"""
Consolidate common headers across all tabs of an Excel workbook.

Logic:
  1. For each sheet, auto-detect the header row (the row with the most
     non-empty string cells — beats trying to hard-code row 13).
  2. Read data rows below the header.
  3. Take the intersection of headers across all sheets = "common headers".
  4. Write a "Consolidated" tab: [Source Sheet] + common headers + data.
"""

from collections import Counter
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

SRC = Path(__file__).parent / "Sample2.xlsx"
OUT = Path(__file__).parent / "Sample2_Consolidated.xlsx"

# STRICT: only headers present in EVERY sheet get consolidated.
# THRESHOLD (e.g. 0.5): headers present in >= 50% of sheets get consolidated;
#                       sheets missing a header just leave blanks.
MODE = "THRESHOLD"
THRESHOLD = 0.5
CONSOLIDATED_TAB = "Consolidated"


def detect_header_row(ws, scan_rows=40):
    """Return 1-indexed row number that looks most like a header row."""
    best_row, best_score = 1, -1
    for r in range(1, min(ws.max_row, scan_rows) + 1):
        vals = [c.value for c in ws[r]]
        score = sum(1 for v in vals if isinstance(v, str) and v.strip())
        if score > best_score:
            best_score, best_row = score, r
    return best_row


def read_sheet(ws):
    """Return (headers, list-of-row-dicts) for a sheet.

    Generic 'Quantity' columns are qualified with the preceding descriptive
    header so they stay unique (e.g. '1.1 Pre Carriage ... — Quantity').
    """
    hdr_row = detect_header_row(ws)
    raw = [
        (str(c.value).strip() if c.value is not None else "")
        for c in ws[hdr_row]
    ]
    headers, last_desc = [], ""
    for h in raw:
        if h and h.lower() != "quantity":
            last_desc = h
            headers.append(h)
        elif h.lower() == "quantity":
            headers.append(f"{last_desc} — Quantity" if last_desc else "Quantity")
        else:
            headers.append("")
    rows = []
    for r in ws.iter_rows(min_row=hdr_row + 1, values_only=True):
        if all(v is None or v == "" for v in r):
            continue
        rows.append({h: v for h, v in zip(headers, r) if h})
    return headers, rows


def main():
    wb = openpyxl.load_workbook(SRC, data_only=True)

    # 1. Read every sheet (skip the hidden Formulas_* sheet).
    per_sheet = {}
    for name in wb.sheetnames:
        if name.lower().startswith("formulas_"):
            continue
        headers, rows = read_sheet(wb[name])
        if headers and rows:
            per_sheet[name] = (headers, rows)

    if not per_sheet:
        print("No data sheets found.")
        return

    # 2. Score every header by how many sheets it appears in.
    n_sheets = len(per_sheet)
    hdr_freq = Counter()
    for hs, _ in per_sheet.values():
        hdr_freq.update(set(h for h in hs if h))

    if MODE == "STRICT":
        keep = {h for h, n in hdr_freq.items() if n == n_sheets}
    else:
        keep = {h for h, n in hdr_freq.items() if n / n_sheets >= THRESHOLD}

    # 3. Preserve column order: take union order from the first sheet, then
    #    append any leftover kept headers seen in later sheets.
    ordered_common, seen = [], set()
    for hs, _ in per_sheet.values():
        for h in hs:
            if h in keep and h not in seen:
                seen.add(h)
                ordered_common.append(h)

    print(f"Sheets consolidated : {n_sheets}   Mode: {MODE}"
          + (f" (>= {int(THRESHOLD*100)}%)" if MODE == "THRESHOLD" else ""))
    print(f"Consolidated headers ({len(ordered_common)}):")
    for h in ordered_common:
        print(f"   [{hdr_freq[h]:>2}/{n_sheets}]  {h}")

    dropped = [h for h in hdr_freq if h not in keep]
    if dropped:
        print(f"\nDropped ({len(dropped)} headers below threshold) — sample:")
        for h in sorted(dropped, key=lambda x: -hdr_freq[x])[:10]:
            print(f"   [{hdr_freq[h]:>2}/{n_sheets}]  {h}")

    # 5. Build the consolidated sheet.
    if CONSOLIDATED_TAB in wb.sheetnames:
        del wb[CONSOLIDATED_TAB]
    out = wb.create_sheet(CONSOLIDATED_TAB, 0)

    final_headers = ["Source Sheet"] + ordered_common
    out.append(final_headers)

    bold = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="305496")
    for cell in out[1]:
        cell.font = bold
        cell.fill = fill

    for sheet_name, (_, rows) in per_sheet.items():
        for row in rows:
            out.append([sheet_name] + [row.get(h, "") for h in ordered_common])

    # Freeze header + autosize.
    out.freeze_panes = "A2"
    for i, h in enumerate(final_headers, 1):
        width = max(len(str(h)), 15)
        out.column_dimensions[openpyxl.utils.get_column_letter(i)].width = min(width + 2, 40)

    try:
        wb.save(OUT)
    except PermissionError:
        alt = OUT.with_stem(OUT.stem + "_new")
        wb.save(alt)
        print(f"\n{OUT.name} is open in Excel — wrote {alt.name} instead.")
        return
    print(f"\nWritten -> {OUT.name}  ({out.max_row - 1} rows)")


if __name__ == "__main__":
    main()
