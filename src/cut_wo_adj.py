"""Cut the "- wo adj" copies from a period that has been signed off.

These are snapshots handed to SCM, not another revision of the forecast: P110 and P120
are copied unchanged, and P110A additionally has the manual adjustment row blanked
(销售收入_China!D672:O672, configured under wo_adj.blank).

Two things follow from "snapshot":

  - It runs after sign-off, never before. Cutting early captures numbers that are still
    moving, and the SCM files downstream have no way to tell.
  - Links are not updated when the copy is opened. Pulling fresh values out of the link
    sources would make it a different workbook from the one that was signed off.

Blanking has to go through Excel. Deleting the cells straight out of the package is
quick, but every formula that sums that row keeps its cached total, so the adjustment
would still be sitting in the numbers SCM reads.

    python src/cut_wo_adj.py --period 7+5            # plan only
    python src/cut_wo_adj.py --period 7+5 --apply
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

XL_CALCULATION_AUTOMATIC = -4105
XL_DONE = 0
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


def plan(cfg: dict, period: str) -> list[tuple[str, Path, Path, dict | None]]:
    """(id, official workbook, wo adj copy, range to blank) for each workbook."""
    wo = cfg.get("wo_adj", {})
    suffix = wo.get("suffix", " - wo adj")
    blanks = {b["workbook"]: b for b in wo.get("blank", [])}
    base = Path(cfg["period_root"]) / period
    out = []
    for wbcfg in sorted(cfg["workbooks"], key=lambda w: w["order"]):
        src = base / wbcfg["relative_dir"] / wbcfg["name_template"].format(period=period)
        dst = src.with_name(src.stem + suffix + src.suffix)
        out.append((wbcfg["id"], src, dst, blanks.get(wbcfg["id"])))
    return out


def flatten(value) -> list:
    """Excel hands back a scalar, a tuple, or a tuple of row tuples."""
    if not isinstance(value, tuple):
        return [value]
    out = []
    for item in value:
        out.extend(flatten(item))
    return out


def blank_range(path: Path, sheet: str, cell_range: str, timeout_s: int = 900) -> list:
    """Clear a range and recalculate, returning what was in it."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    try:
        app.Visible = False
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        app.EnableEvents = False
        app.AskToUpdateLinks = False
        # These workbooks carry macros that have no business running here.
        app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE

        wb = app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
        try:
            app.Calculation = XL_CALCULATION_AUTOMATIC
            names = [ws.Name for ws in wb.Worksheets]
            if sheet not in names:
                raise KeyError(f"sheet {sheet!r} not found; workbook has {names}")

            rng = wb.Worksheets(sheet).Range(cell_range)
            before = flatten(rng.Value)
            rng.ClearContents()

            app.CalculateFullRebuild()
            deadline = time.time() + timeout_s
            while app.CalculationState != XL_DONE:
                if time.time() > deadline:
                    raise TimeoutError(f"recalculation exceeded {timeout_s}s")
                time.sleep(0.5)

            wb.Save()
            return before
        finally:
            wb.Close(SaveChanges=False)
    finally:
        app.Quit()
        pythoncom.CoUninitialize()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--period", required=True, help="the period that has been signed off")
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--force", action="store_true", help="overwrite existing wo adj copies")
    ap.add_argument("--config", default=str(ROOT / "config" / "roll.json"))
    args = ap.parse_args()

    cfg = json.loads(Path(args.config).read_text(encoding="utf-8"))
    items = plan(cfg, args.period)

    blocked = False
    for wb_id, src, dst, spec in items:
        note = f"blank {spec['sheet']}!{spec['range']}" if spec else "straight copy"
        print(f"  {wb_id}: {dst.name}\n      {note}")
        if not src.exists():
            print(f"      SOURCE MISSING: {src}")
            blocked = True
        elif dst.exists() and not args.force:
            print("      already exists - re-run with --force to replace")
            blocked = True

    if not args.apply:
        print("\nplan only - nothing written. add --apply to execute.")
        return
    if blocked:
        sys.exit("\nrefusing to write until the problems above are resolved")

    print()
    for wb_id, src, dst, spec in items:
        shutil.copy2(src, dst)
        print(f"  {wb_id}: copied -> {dst.name}", flush=True)
        if not spec:
            continue
        print(f"      opening Excel to blank {spec['sheet']}!{spec['range']}", flush=True)
        before = blank_range(dst, spec["sheet"], spec["range"])
        kept = [v for v in before if v not in (None, "")]
        print(f"      cleared {len(kept)} non-empty cell(s): {kept if kept else '(row was empty)'}")

    print(f"\ndone. next: python src/scm.py --from <上期> --to {args.period} --apply")


if __name__ == "__main__":
    main()
