"""Print the stored values of a sheet region, to inspect layout without opening Excel."""

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from compare_cache import live_values

CELL_RE = re.compile(r"([A-Z]+)(\d+)")


def col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("--sheet", action="append", required=True)
    ap.add_argument("--max-row", type=int, default=12)
    ap.add_argument("--max-col", type=int, default=10)
    args = ap.parse_args()

    data = live_values(args.workbook)
    for name in args.sheet:
        match = next((s for s in data if s.lower() == name.lower()), None)
        if match is None:
            print(f"\n### {name}: NOT FOUND (available: {sorted(data)})")
            continue
        cells = data[match]
        print(f"\n### {match}  ({len(cells)} non-empty cells)")
        rows: dict[int, dict[int, str]] = {}
        for ref, (_, val) in cells.items():
            m = CELL_RE.fullmatch(ref)
            if not m:
                continue
            c, r = col_index(m.group(1)), int(m.group(2))
            if r <= args.max_row and c <= args.max_col:
                rows.setdefault(r, {})[c] = str(val)[:28]
        for r in sorted(rows):
            cols = " | ".join(f"{c}:{v}" for c, v in sorted(rows[r].items()))
            print(f"  r{r}: {cols}")


if __name__ == "__main__":
    main()
