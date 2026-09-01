"""Run the whole monthly roll: copy the period folder, re-point links, recalculate.

    python src/run_month.py --from 6+6 --to 7+5              # plan only, writes nothing
    python src/run_month.py --from 6+6 --to 7+5 --apply      # copy + re-point
    python src/run_month.py --from 6+6 --to 7+5 --apply --refresh   # also recalculate

Stops before recalculating if any link is unresolved, so a bad plan cannot quietly
produce a workbook full of stale numbers.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path, PureWindowsPath

sys.path.insert(0, str(Path(__file__).resolve().parent))

from netfs import NetFS
from relink import (REPAIRED, ROLLED, UNRESOLVED, WAITING, Rewriter, apply_plan,
                    build_plan)

ROOT = Path(__file__).resolve().parent.parent


def workbook_paths(cfg: dict, period: str) -> list[tuple[str, Path]]:
    """Core workbooks for a period, in dependency order."""
    base = Path(cfg["period_root"]) / period
    out = []
    for wbcfg in sorted(cfg["workbooks"], key=lambda w: w["order"]):
        name = wbcfg["name_template"].format(period=period)
        out.append((wbcfg["id"], base / wbcfg["relative_dir"] / name))
    return out


def roll_in_place(path: Path, rw: Rewriter, fs: NetFS, cfg: dict, apply: bool):
    plans = build_plan(str(path), rw, fs, base_dir=str(path.parent))
    counts: dict[str, int] = {}
    for p in plans:
        counts[p.status] = counts.get(p.status, 0) + 1
        if p.new_abs != p.old_abs:
            print(f"    [{p.link_no}] {p.status}: {PureWindowsPath(p.new_abs).name}")
            if p.note:
                print(f"          {p.note}")
    print(f"    summary: " + ", ".join(f"{k}={v}" for k, v in sorted(counts.items())))

    unresolved = [p for p in plans if p.status == UNRESOLVED]
    for p in unresolved:
        print(f"    UNRESOLVED [{p.link_no}] {p.new_abs}")
    for p in (p for p in plans if p.status == WAITING):
        print(f"    WAITING    [{p.link_no}] {PureWindowsPath(p.new_abs).name}")

    if apply and not unresolved:
        changed = [p for p in plans if p.status in (ROLLED, REPAIRED, WAITING)
                   and p.new_abs != p.old_abs]
        if changed:
            # Written beside the original: a temp file on the local disk cannot be
            # replaced onto the share, they are different volumes.
            tmp = path.with_suffix(path.suffix + ".tmp")
            apply_plan(str(path), str(tmp), plans, cfg["share_prefix"], base_dir=str(path.parent))
            tmp.replace(path)
            print(f"    rewrote {len(changed)} link(s) in place")
    return plans


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--from", dest="from_period", required=True)
    ap.add_argument("--to", dest="to_period", required=True)
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--refresh", action="store_true", help="also recalculate in Excel")
    ap.add_argument("--include-feedback", action="store_true")
    ap.add_argument("--skip-hfm-check", action="store_true",
                    help="recalculate even if the Smart View grids look stale")
    ap.add_argument("--config", default=str(ROOT / "config" / "roll.json"))
    args = ap.parse_args()

    cfg = json.loads(Path(args.config).read_text(encoding="utf-8"))

    print("=" * 70, flush=True)
    print(f"STEP 1  copy period folder {args.from_period} -> {args.to_period}", flush=True)
    print("=" * 70, flush=True)
    cmd = [sys.executable, str(ROOT / "src" / "copy_period.py"),
           "--from", args.from_period, "--to", args.to_period, "--config", args.config]
    if args.apply:
        cmd.append("--apply")
    if args.include_feedback:
        cmd.append("--include-feedback")
    if subprocess.call(cmd) != 0:
        sys.exit("folder copy failed")

    if not args.apply:
        print("\nplan only - stopping before link rewrite. add --apply to continue.")
        return

    print("\n" + "=" * 70)
    print("STEP 2  re-point external links (P120 -> P110A -> P110)")
    print("=" * 70)

    rw = Rewriter(
        share_prefix=cfg["share_prefix"],
        managed_root=cfg["managed_root"],
        period_root=cfg["period_root"],
        from_period=args.from_period,
        to_period=args.to_period,
        fiscal_year=cfg.get("fiscal_year", 2026),
        months=cfg.get("months_in_year", 12),
        pinned_periods=set(cfg.get("pinned_periods", [])),
    )
    fs = NetFS(cache_file=str(ROOT / ".cache" / "dirs.json"))
    fs.invalidate(cfg["period_root"])

    targets = workbook_paths(cfg, args.to_period)
    blocked = False
    for wb_id, path in targets:
        print(f"\n  {wb_id}: {path.name}")
        if not path.exists():
            print("    NOT FOUND - was the folder copy run?")
            blocked = True
            continue
        fs.invalidate(str(path.parent))
        plans = roll_in_place(path, rw, fs, cfg, apply=True)
        if any(p.status == UNRESOLVED for p in plans):
            blocked = True
    fs.save()

    if blocked:
        sys.exit("\nunresolved links remain - fix them before recalculating")

    if not args.refresh:
        print("\nlinks updated. add --refresh to recalculate in Excel.")
        return

    print("\n" + "=" * 70)
    print("STEP 3  check the Smart View grids were refreshed")
    print("=" * 70)
    from hfm_guard import check

    actual_months = int(args.to_period.split("+")[0])
    stale: list[str] = []
    for wb_id, path in targets:
        for sheet, _, verdict in check(str(path), cfg.get("fiscal_year", 2026), actual_months):
            if verdict == "not in use":
                continue
            print(f"  {wb_id} / {sheet}: {verdict}")
            if not verdict.startswith("ok"):
                stale.append(f"{wb_id} / {sheet}")

    if stale and not args.skip_hfm_check:
        print("\nThese grids still hold an earlier month's actuals:")
        for s in stale:
            print(f"  {s}")
        sys.exit("refresh them in Smart View, then re-run. "
                 "(--skip-hfm-check overrides, but every reconciliation will be off.)")

    print("\n" + "=" * 70)
    print("STEP 4  recalculate in Excel")
    print("=" * 70)
    from excel_refresh import refresh
    refresh([str(p) for _, p in targets])
    print("\ndone. verify with tools/compare_workbooks.py before sending out.")
    print("next: cut the '- wo adj' copies from the signed-off workbooks (blank "
          "销售收入_China D672:O672), then hand the result to SCM with\n"
          f"  python src/scm.py --from {args.from_period} --to {args.to_period} --apply")


if __name__ == "__main__":
    main()
