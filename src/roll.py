"""Preview or apply the external-link roll for one workbook.

Preview:  python src/roll.py --workbook <path> --from 5+7 --to 6+6 [--base-dir <share folder>]
Apply:    python src/roll.py --workbook <path> --from 5+7 --to 6+6 --apply <out path>
Validate: add --compare-to <workbook> to diff the plan against a hand-made result.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path, PureWindowsPath
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parent))

from netfs import NetFS
from relink import (NS_PKG_REL, REPAIRED, ROLLED, UNRESOLVED, WAITING, Rewriter,
                    apply_plan, build_plan)

ROOT = Path(__file__).resolve().parent.parent


def period_sequence(months: int) -> list[str]:
    return [f"{n}+{months - n}" for n in range(months + 1)]


def read_link_targets(workbook: str, rw: Rewriter, base_dir: str) -> dict[int, str]:
    """Absolute target of each external link in an existing workbook."""
    wb_dir = PureWindowsPath(base_dir)
    out: dict[int, str] = {}
    with zipfile.ZipFile(workbook) as zf:
        for name in zf.namelist():
            m = re.fullmatch(r"xl/externalLinks/_rels/externalLink(\d+)\.xml\.rels", name)
            if not m:
                continue
            rels = ET.fromstring(zf.read(name)).findall(f"{{{NS_PKG_REL}}}Relationship")
            targets = [rw.to_absolute(r.get("Target") or "", wb_dir) for r in rels]
            if targets:
                out[int(m.group(1))] = max(targets, key=len)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook", required=True)
    ap.add_argument("--from", dest="from_period", required=True)
    ap.add_argument("--to", dest="to_period", required=True)
    ap.add_argument("--apply", dest="out_path")
    ap.add_argument("--base-dir", help="share folder the workbook logically lives in")
    ap.add_argument("--compare-to", help="existing next-period workbook to validate the plan against")
    ap.add_argument("--compare-base-dir")
    ap.add_argument("--absolute-links", action="store_true",
                    help="write every target as an absolute path (use when the output "
                         "will not sit in its normal folder)")
    ap.add_argument("--config", default=str(ROOT / "config" / "roll.json"))
    args = ap.parse_args()

    cfg = json.loads(Path(args.config).read_text(encoding="utf-8"))
    seq = period_sequence(cfg.get("months_in_year", 12))
    for p in (args.from_period, args.to_period):
        if p not in seq:
            sys.exit(f"period {p!r} must be one of {seq}")

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
    plans = build_plan(args.workbook, rw, fs, base_dir=args.base_dir)

    counts: dict[str, int] = {}
    print(f"{'link':>5}  {'status':<10}  target")
    for p in sorted(plans, key=lambda x: x.link_no):
        counts[p.status] = counts.get(p.status, 0) + 1
        print(f"[{p.link_no:>3}]  {p.status:<10}  {p.new_abs}")
        if p.new_abs != p.old_abs:
            print(f"{'':>5}  {'':<10}  was: {p.old_abs}")
        if p.note:
            print(f"{'':>5}  {'':<10}  note: {p.note}")
    print("\nsummary: " + ", ".join(f"{k}={v}" for k, v in sorted(counts.items())))

    if args.compare_to:
        actual = read_link_targets(args.compare_to, rw, args.compare_base_dir or args.base_dir)
        planned = {p.new_abs.lower() for p in plans}
        actual_set = {v.lower() for v in actual.values()}
        print("\n=== PLAN vs HAND-MADE RESULT ===")
        print(f"targets in both        : {len(planned & actual_set)}")
        only_plan = sorted(planned - actual_set)
        only_actual = sorted(actual_set - planned)
        for label, items in (("only in plan", only_plan), ("only in hand-made file", only_actual)):
            print(f"{label:<23}: {len(items)}")
            for i in items:
                print(f"    {i}")

    fs.save()

    waiting = [p for p in plans if p.status == WAITING]
    if waiting:
        print(f"\n{len(waiting)} link(s) point at inputs that have not arrived yet:")
        for p in waiting:
            print(f"  [{p.link_no}] {p.new_abs}")

    unresolved = [p for p in plans if p.status == UNRESOLVED]
    if args.out_path:
        if unresolved:
            sys.exit(f"\nrefusing to write: {len(unresolved)} unresolved link(s)")
        n = apply_plan(args.workbook, args.out_path, plans, cfg["share_prefix"],
                       base_dir=args.base_dir, force_absolute=args.absolute_links)
        changed = sum(1 for p in plans if p.status in (ROLLED, REPAIRED))
        print(f"\nwrote {args.out_path} ({n} entries rewritten across {changed} links)")


if __name__ == "__main__":
    main()
