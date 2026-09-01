"""Roll the P120 Cover FX formulas for the target forecast period.

The external-link rewrite moves P120 to the new monthly Rate workbook, but the
Cover sheet also carries month-specific formulas:

* actual months use GC Avg Rate
* the first forecast month uses GC Corporate Rate from the latest actual month
* later forecast months equal the previous month

This script writes those formulas through Excel COM so Excel keeps its shared
formula structures and cached values healthy.
"""

from __future__ import annotations

import argparse
import calendar
import json
import re
import sys
from pathlib import Path, PureWindowsPath

ROOT = Path(__file__).resolve().parent.parent


def col_index(col: str) -> int:
    out = 0
    for ch in col:
        out = out * 26 + ord(ch.upper()) - 64
    return out


def col_name(idx: int) -> str:
    out = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        out = chr(65 + rem) + out
    return out


def add_cols(col: str, offset: int) -> str:
    return col_name(col_index(col) + offset)


def actual_month(period: str) -> int:
    match = re.fullmatch(r"(\d+)\+(\d+)", period)
    if not match:
        raise ValueError(f"period must look like '7+5', got {period!r}")
    month = int(match.group(1))
    if not 1 <= month <= 12:
        raise ValueError("P120 FX formula roll expects a period from 1+11 through 12+0")
    return month


def default_workbook(cfg: dict, period: str) -> Path:
    return (
        Path(cfg["period_root"])
        / period
        / "P1"
        / f"P120 Interco sales_China CY26 {period} FCST.xlsx"
    )


def rate_workbook(cfg: dict, month: int) -> Path:
    year = cfg.get("fiscal_year", 2026)
    last_day = calendar.monthrange(year, month)[1]
    root = Path(cfg["period_root"]).parent.parent
    return root / "FX Rate" / f"Rate {year}{month:02d}{last_day:02d}.xlsx"


def formula_prefix(rate_path: Path) -> str:
    folder = str(PureWindowsPath(rate_path.parent))
    return f"='{folder}\\[{rate_path.name}]"


def planned_formulas(cfg: dict, period: str) -> dict[str, str]:
    month = actual_month(period)
    prefix = formula_prefix(rate_workbook(cfg, month))
    rows = {16: 5, 18: 14}  # Cover row -> Rate workbook currency row
    formulas: dict[str, str] = {}

    for cover_row, rate_row in rows.items():
        for m in range(1, 13):
            cover_cell = f"{add_cols('D', m - 1)}{cover_row}"
            if m <= month:
                avg_col = add_cols("Y", m - 1)
                formulas[cover_cell] = f"{prefix}GC Avg Rate'!{avg_col}${rate_row}"
            elif m == month + 1:
                corp_col = add_cols("AO", month - 1)
                formulas[cover_cell] = f"{prefix}GC Corporate Rate'!{corp_col}${rate_row}"
            else:
                prev_cell = f"{add_cols('D', m - 2)}{cover_row}"
                formulas[cover_cell] = f"={prev_cell}"
    return formulas


def write_formulas(workbook: Path, formulas: dict[str, str], apply: bool, visible: bool) -> None:
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = visible
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.AskToUpdateLinks = False
    try:
        wb = excel.Workbooks.Open(str(workbook), UpdateLinks=0, ReadOnly=not apply)
        try:
            ws = wb.Worksheets("Cover")
            for cell, formula in formulas.items():
                if apply:
                    ws.Range(cell).Formula = formula
                print(f"{cell}: {ws.Range(cell).Formula if not apply else formula}")
            if apply:
                excel.CalculateFull()
                wb.Save()
                print(f"\nsaved {workbook}")
            else:
                print("\ndry run - add --apply to write these formulas")
        finally:
            wb.Close(SaveChanges=False)
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--period", required=True, help="target period, e.g. 7+5")
    ap.add_argument("--workbook", help="P120 workbook path; defaults to config path")
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--visible", action="store_true")
    ap.add_argument("--config", default=str(ROOT / "config" / "roll.json"))
    args = ap.parse_args()

    cfg = json.loads(Path(args.config).read_text(encoding="utf-8"))
    workbook = Path(args.workbook) if args.workbook else default_workbook(cfg, args.period)
    if not workbook.exists():
        sys.exit(f"workbook not found: {workbook}")

    rate = rate_workbook(cfg, actual_month(args.period))
    if not rate.exists():
        sys.exit(f"rate workbook not found: {rate}")

    write_formulas(workbook, planned_formulas(cfg, args.period), args.apply, args.visible)


if __name__ == "__main__":
    main()
