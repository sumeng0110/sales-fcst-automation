"""Re-point the SCM handover files at this period's "- wo adj" workbooks.

SCM is the downstream consumer of the forecast, not an input: the numbered files 1-5
under SCM\\CY26 <period> read the three "- wo adj" copies of P110 / P110A / P120, and
feed each other in the order their numbers suggest. Nothing in the forecast chain reads
SCM back, so this step runs last - and only once the wo adj copies exist, otherwise
every target resolves to nothing.

The rule here is broader than the main chain's: a period token advances wherever it
appears in the path, because these files reference three different trees at once. That
also happens to be what makes each file's link to the previous period's own copy come
out right - the folder moves to this period while the file name moves to the last one.

    python src/scm.py --from 6+6 --to 7+5             # plan only
    python src/scm.py --from 6+6 --to 7+5 --apply     # rewrite the 7+5 files in place
    python src/scm.py --from 5+7 --to 6+6 --validate  # check the rules against history
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
from pathlib import Path, PureWindowsPath

sys.path.insert(0, str(Path(__file__).resolve().parent))

from netfs import NetFS
from relink import (MISSING, PINNED, REPAIRED, ROLLED, UNRESOLVED, WAITING, LinkPlan,
                    Rewriter, apply_plan, build_plan, decide, normalize)
from roll import read_link_targets

ROOT = Path(__file__).resolve().parent.parent

# The handover files are numbered to record the order they must be produced in.
NUMBERED_RE = re.compile(r"^[1-5]\.")

# SCM stamps the year with two digits ("预测数据2606  6+6.xlsx") where the forecast
# chain uses four ("202606"), so the main advance_stamp does not see them.
SHORT_STAMP_RE = re.compile(r"(?<!\d)(\d{2})(0[1-9]|1[0-2])(?!\d)")

# PERIOD_RE refuses a token followed by a dot so that it cannot bite into a version
# number, but SCM puts the period right before the extension ("转移比例 CY26 5+7.xlsx").
# Only a dot that starts another number is still excluded.
PERIOD_IN_PATH_RE = re.compile(r"(?<![\d.+])(\d{1,2}\+\d{1,2})(?![\d+]|\.\d)")


def period_folder(cfg: dict, period: str) -> str:
    return str(PureWindowsPath(cfg["scm_root"], f"{cfg['year_token']} {period}"))


def advance_periods(text: str, rw: Rewriter) -> str:
    """Advance every period token in a path by one, wherever it appears.

    Folder and file name move independently, which is what makes each file's reference
    to the previous period's own copy land correctly: the folder becomes this period
    while the name inside it becomes the last one.
    """
    def sub(m):
        return rw.advance(m.group(1)) or m.group(0)
    return PERIOD_IN_PATH_RE.sub(sub, text)


def advance_short_stamp(name: str, rw: Rewriter) -> str:
    """Move a YYMM stamp naming the source month to the target month."""
    def sub(m):
        if (int(m.group(1)), int(m.group(2))) != (rw.fiscal_year % 100, rw.from_month):
            return m.group(0)
        return f"{rw.fiscal_year % 100:02d}{rw.to_month:02d}"
    return SHORT_STAMP_RE.sub(sub, name)


def make_decider(rw: Rewriter, cfg: dict):
    """Decision rules for a handover file, layered on top of the main chain's."""
    scm_root = cfg["scm_root"].rstrip("\\").lower()

    def decide_scm(link_no: int, old_abs: str, _rw, fs, folder_for) -> LinkPlan:
        if not rw.is_managed(old_abs):
            return LinkPlan(link_no, old_abs, old_abs, PINNED,
                            "outside managed root (dead link to a personal drive)")

        # The wo adj copies live in the forecast tree, so the already validated rule
        # applies unchanged - including its repair of a file name left on the old period.
        if rw.period_segment(old_abs) is not None:
            return decide(link_no, old_abs, rw, fs, folder_for)

        p = PureWindowsPath(old_abs)
        perioded_name = advance_periods(p.name, rw)
        if "预测数据" in p.name:
            # 预测数据 is prepared by hand for the target period but can still carry
            # the latest actual month in its YYMM stamp, e.g. "预测数据2606  7+5.xlsx".
            # Because these are often relative links, the copied workbook already
            # resolves the parent to the target SCM folder; pin the parent there and
            # only roll the period token in the file name.
            name = perioded_name
            new_abs = normalize(str(PureWindowsPath(period_folder(cfg, rw.to_period), name)))
        else:
            name = advance_short_stamp(perioded_name, rw)
            new_abs = normalize(str(PureWindowsPath(advance_periods(str(p.parent), rw), name)))
        step = f"{rw.from_period} -> {rw.to_period}"

        # A numbered handover file lives in the folder of the period it names. When a
        # copy of one gets left behind in a later folder and something links to that
        # copy, the link is put back on the original instead of carrying the stray
        # copy forward for the rest of the year.
        named = PERIOD_IN_PATH_RE.search(name)
        if named and NUMBERED_RE.match(name) and old_abs.lower().startswith(scm_root + "\\"):
            home = normalize(str(PureWindowsPath(
                cfg["scm_root"], f"{cfg['year_token']} {named.group(1)}", name)))
            if home != new_abs:
                new_abs = home
                step += ", back to the folder it belongs in"

        if new_abs == old_abs:
            status = PINNED if fs.exists(old_abs) else MISSING
            return LinkPlan(link_no, old_abs, old_abs, status, "nothing in the path to advance")
        if fs.exists(new_abs):
            return LinkPlan(link_no, old_abs, new_abs, ROLLED, step)
        if old_abs.lower().startswith(scm_root + "\\"):
            return LinkPlan(link_no, old_abs, new_abs, WAITING,
                            "not in the SCM folder yet; re-run once it is dropped in")
        # Another department's tree. It does advance, but its version subfolder is named
        # by hand (v2, "2+10 - v3"), so a target that is not there is not guessed at.
        return LinkPlan(link_no, old_abs, new_abs, UNRESOLVED,
                        "target missing - check the version subfolder of that period")

    return decide_scm


SKIP_NAMES = {"thumbs.db"}


def folder_plan(cfg: dict, rw: Rewriter) -> list[tuple[Path, str, str]]:
    """(source file, name in the new folder, action) for building the next period.

    Two families behave differently. The numbered handover files and the named-account
    workbook are this period's work, so they move on with the period. 转移比例 and
    预测数据 are a historical series that piles up in every folder, so they keep their
    names - this month's copies are dropped in later by hand, not produced here.
    """
    rename_res = [re.compile(p) for p in cfg.get("scm_folder", {}).get("rename_patterns", [])]
    src = period_folder(cfg, rw.from_period)
    out: list[tuple[Path, str, str]] = []
    try:
        entries = sorted(os.scandir(src), key=lambda e: e.name)
    except OSError as exc:
        raise SystemExit(f"cannot read {src}: {exc}")

    for e in entries:
        if e.is_dir() or e.name.startswith("~$") or e.name.lower() in SKIP_NAMES:
            continue
        if not any(r.search(e.name) for r in rename_res):
            out.append((Path(e.path), e.name, "carry"))
        elif rw.from_period in e.name:
            out.append((Path(e.path), advance_periods(e.name, rw), "rename"))
        else:
            # An earlier period's copy left in this folder. Referencing last period is
            # done across folders, so there is nothing to carry forward.
            out.append((Path(e.path), e.name, "drop"))
    return out


def handover_files(folder: str, period: str) -> list[Path]:
    """The numbered files of one period, in execution order."""
    try:
        entries = os.scandir(folder)
    except OSError:
        return []
    with entries as it:
        found = [Path(e.path) for e in it
                 if e.is_file() and NUMBERED_RE.match(e.name) and period in e.name]
    return sorted(found, key=lambda p: p.name)


def create_folder(cfg: dict, rw: Rewriter, apply: bool, validate: bool) -> int:
    """Build the next period's SCM folder, or check the rules against a real one."""
    items = folder_plan(cfg, rw)
    dst = Path(period_folder(cfg, rw.to_period))
    kept = [(s, n, a) for s, n, a in items if a != "drop"]

    print(f"{period_folder(cfg, rw.from_period)}\n  -> {dst}\n")
    for _, name, action in items:
        print(f"  {action:<7} {name}")
    print(f"\n{len(kept)} file(s) in, {len(items) - len(kept)} dropped")

    if validate:
        try:
            actual = {e.name for e in os.scandir(dst) if e.is_file()
                      and not e.name.startswith("~$") and e.name.lower() not in SKIP_NAMES}
        except OSError as exc:
            print(f"\ncannot read {dst}: {exc}")
            return 1
        planned = {n for _, n, a in items if a != "drop"}
        print(f"\nvs the real {rw.to_period} folder: {len(planned & actual)} of "
              f"{len(actual)} names match")
        for n in sorted(planned - actual):
            print(f"  only in plan       : {n}")
        for n in sorted(actual - planned):
            print(f"  only in real folder: {n}")
        return 0 if planned == actual else 1

    if not apply:
        print("\nplan only - nothing written. add --apply to execute.")
        return 0
    if dst.exists():
        print(f"\n{dst} already exists; existing files are left alone")
    dst.mkdir(parents=True, exist_ok=True)
    written = 0
    for src, name, action in kept:
        target = dst / name
        if target.exists():
            continue
        shutil.copy2(src, target)
        written += 1
        print(f"  wrote {name}", flush=True)
    print(f"\ncopied {written} file(s). next: cut the wo adj copies, then re-run "
          f"with --apply to re-point the links.")
    return 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--from", dest="from_period", required=True)
    ap.add_argument("--to", dest="to_period", required=True)
    ap.add_argument("--apply", action="store_true", help="rewrite the target period's files in place")
    ap.add_argument("--create", action="store_true",
                    help="build the target period's folder instead of re-pointing links")
    ap.add_argument("--validate", action="store_true",
                    help="plan from the source period and diff against the real target files")
    ap.add_argument("--config", default=str(ROOT / "config" / "roll.json"))
    args = ap.parse_args()

    cfg = json.loads(Path(args.config).read_text(encoding="utf-8"))
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
    decider = make_decider(rw, cfg)

    if args.create:
        sys.exit(create_folder(cfg, rw, apply=args.apply, validate=args.validate))

    # Validation reads the source period, because the point is to reproduce a month
    # that was already done by hand. A real run works on the copies already renamed
    # into the new folder, the same way run_month.py does.
    src_folder = period_folder(cfg, args.from_period)
    dst_folder = period_folder(cfg, args.to_period)
    work_folder = src_folder if args.validate else dst_folder
    work_period = args.from_period if args.validate else args.to_period

    files = handover_files(work_folder, work_period)
    print(f"{work_folder}\n{len(files)} handover file(s) carrying '{work_period}'\n")
    if not files:
        sys.exit("nothing to do - has the period folder been created and the files renamed?")

    counterparts = handover_files(dst_folder, args.to_period) if args.validate else []

    exit_code = 0
    for path in files:
        print(f"  {path.name}")
        plans = build_plan(str(path), rw, fs, base_dir=str(path.parent), decide_fn=decider)
        for p in sorted(plans, key=lambda x: x.link_no):
            arrow = PureWindowsPath(p.new_abs).name if p.new_abs != p.old_abs else "(unchanged)"
            print(f"    [{p.link_no}] {p.status:<10} {arrow}")
            if p.note and p.status not in (ROLLED, PINNED):
                print(f"          {p.note}")

        if args.validate:
            actual_path = next((q for q in counterparts if q.name[:2] == path.name[:2]), None)
            if actual_path is None:
                print(f"    no hand-made counterpart in {args.to_period}\n")
                continue
            actual = {t.lower() for t in
                      read_link_targets(str(actual_path), rw, str(actual_path.parent)).values()}
            planned = {p.new_abs.lower() for p in plans}
            same = planned & actual
            print(f"    vs hand-made: {len(same)}/{len(actual)} targets match")
            for extra in sorted(planned - actual):
                print(f"      only in plan  : {extra}")
            for extra in sorted(actual - planned):
                print(f"      only by hand  : {extra}")
            if planned != actual:
                exit_code = 1

        if args.apply and not args.validate:
            unresolved = [p for p in plans if p.status == UNRESOLVED]
            if unresolved:
                print(f"    refusing to write: {len(unresolved)} unresolved link(s)")
                exit_code = 1
                continue
            changed = [p for p in plans if p.status in (ROLLED, REPAIRED, WAITING)
                       and p.new_abs != p.old_abs]
            if changed:
                # Written beside the original: a temp file on the local disk cannot be
                # os.replace'd onto the share, they are different volumes.
                tmp = path.with_suffix(path.suffix + ".tmp")
                apply_plan(str(path), str(tmp), plans, cfg["share_prefix"],
                           base_dir=str(path.parent))
                tmp.replace(path)
            print(f"    rewrote {len(changed)} link(s)")
        print()

    fs.save()
    if not args.apply and not args.validate:
        print("plan only - nothing written. add --apply to execute.")
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
