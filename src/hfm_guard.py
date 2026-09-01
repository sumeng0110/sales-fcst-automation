"""Check that the Smart View ad-hoc grids were refreshed for the current period.

The HFM sheets hold actuals retrieved by hand through Smart View. Nothing in the
workbook fails if that refresh is skipped - the grid simply keeps last month's numbers
and every downstream check quietly reconciles against the wrong actuals. This module
reads the grids and reports which months actually carry data.

Grid layout: a row of period headers (P01..P12), usually a row of years above it, and
data underneath. Column and row offsets differ between sheets, so both header rows are
located by content rather than by fixed position.
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))

from compare_cache import live_values

CELL_RE = re.compile(r"([A-Z]+)(\d+)")
PERIOD_HEADER_RE = re.compile(r"^P(0[1-9]|1[0-2])$")
YEAR_RE = re.compile(r"^(20\d{2})$")
GRID_NAME_RE = re.compile(r"hfm|adhoc", re.IGNORECASE)
INVALID_MARKERS = {"#invalid", "#missing", "#error", "#ref!", "#value!"}


@dataclass
class MonthColumn:
    period: int
    year: int | None
    populated: int
    invalid: int


def _col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _as_grid(cells: dict[str, tuple[str, str]]) -> dict[int, dict[int, tuple[str, str]]]:
    grid: dict[int, dict[int, tuple[str, str]]] = {}
    for ref, val in cells.items():
        m = CELL_RE.fullmatch(ref)
        if m:
            grid.setdefault(int(m.group(2)), {})[_col_index(m.group(1))] = val
    return grid


def scan_sheet(cells: dict[str, tuple[str, str]]) -> list[MonthColumn] | None:
    """Locate the period header row and summarise the data under each month column."""
    grid = _as_grid(cells)
    header_row = None
    for r in sorted(grid)[:20]:
        periods = {c: PERIOD_HEADER_RE.match(str(v[1]).strip())
                   for c, v in grid[r].items()}
        if sum(1 for m in periods.values() if m) >= 3:
            header_row = r
            break
    if header_row is None:
        return None

    months = {c: int(m.group(1)) for c, v in grid[header_row].items()
              if (m := PERIOD_HEADER_RE.match(str(v[1]).strip()))}

    # These grids often carry an extra column to the right reusing the P12 header for
    # a full-year total. Keeping the leftmost column per period keeps the monthly
    # sequence honest, so a stale P06 is not masked by a populated total.
    leftmost: dict[int, int] = {}
    for col, period in sorted(months.items()):
        leftmost.setdefault(period, col)
    months = {col: period for period, col in leftmost.items()}

    years: dict[int, int] = {}
    for r in range(header_row - 1, max(header_row - 6, 0), -1):
        for c, v in grid.get(r, {}).items():
            if c in months and c not in years and (m := YEAR_RE.match(str(v[1]).strip())):
                years[c] = int(m.group(1))
        if len(years) == len(months):
            break

    out = []
    for c, period in sorted(months.items(), key=lambda kv: kv[1]):
        populated = invalid = 0
        for r, row in grid.items():
            if r <= header_row or c not in row:
                continue
            raw = str(row[c][1]).strip()
            if raw.lower() in INVALID_MARKERS:
                invalid += 1
                continue
            try:
                if float(raw) != 0.0:
                    populated += 1
            except ValueError:
                continue
        out.append(MonthColumn(period, years.get(c), populated, invalid))
    return out


def check(workbook: str, fiscal_year: int, actual_months: int):
    """Report every ad-hoc grid and whether the current period's actuals are present."""
    data = live_values(workbook)
    findings = []
    for sheet in sorted(data):
        if not GRID_NAME_RE.search(sheet):
            continue
        cols = scan_sheet(data[sheet])
        if cols is None:
            findings.append((sheet, None, "no period header row found"))
            continue
        this_year = [c for c in cols if c.year == fiscal_year and c.populated]
        if not this_year:
            # Abandoned grids (all #Invalid, or only prior-year columns) are not part
            # of the monthly routine and would otherwise cry wolf every run.
            findings.append((sheet, cols, "not in use"))
            continue
        last = max(c.period for c in this_year)
        if last < actual_months:
            findings.append((sheet, cols, f"last month with data is P{last:02d}, "
                                          f"expected at least P{actual_months:02d}"))
        else:
            findings.append((sheet, cols, f"ok (data through P{last:02d})"))
    return findings


def main():
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("--period", required=True, help="target period, e.g. 7+5")
    ap.add_argument("--fiscal-year", type=int, default=2026)
    ap.add_argument("--offset", type=int, default=0,
                    help="months the grid is expected to lag the period, e.g. 1 when the "
                         "forecast is built before the month closes in HFM")
    ap.add_argument("--strict", action="store_true", help="exit non-zero if any grid is stale")
    args = ap.parse_args()

    actual_months = int(args.period.split("+")[0]) - args.offset
    findings = check(args.workbook, args.fiscal_year, actual_months)
    if not findings:
        print("no HFM / Adhoc grids found")
        return

    stale = []
    for sheet, cols, verdict in findings:
        print(f"\n### {sheet}: {verdict}")
        if not cols or verdict == "not in use":
            continue
        by_year: dict[int | None, list[str]] = {}
        for c in cols:
            mark = f"P{c.period:02d}"
            if c.invalid:
                mark += f"(#invalid x{c.invalid})"
            elif c.populated:
                mark += f"({c.populated})"
            else:
                mark += "(empty)"
            by_year.setdefault(c.year, []).append(mark)
        for year, marks in sorted(by_year.items(), key=lambda kv: (kv[0] is None, kv[0])):
            print(f"  {year or 'no year'}: " + " ".join(marks))
        if not verdict.startswith("ok"):
            stale.append(sheet)

    print(f"\ntarget period {args.period} -> actuals must cover P01..P{actual_months:02d} of {args.fiscal_year}")
    if stale:
        print(f"NEEDS SMART VIEW REFRESH: {', '.join(stale)}")
    if stale and args.strict:
        sys.exit(1)


if __name__ == "__main__":
    main()
