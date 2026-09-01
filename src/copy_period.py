"""Create the next period's folder from the current one.

Mirrors the folder tree, renaming only the core P110 / P110A / P120 workbooks to the
new period. Everything else keeps its name, because a file called "... 5+7 FCST" that
was never refreshed should not start claiming to be this period's data.

Files matching derived_patterns are left out altogether. The "- wo adj" copies are cut
from this period's finished workbook at the very end of the month, so carrying last
month's forward would put the previous period's numbers behind this period's filename -
and the SCM workbooks that read them would never know.

Defaults to a dry run; nothing is written without --apply.
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
from relink import PERIOD_RE

ROOT = Path(__file__).resolve().parent.parent
TEMP_FILE_RE = re.compile(r"^~\$")


def resolve_source_folder(period_root: str, period: str, fs: NetFS) -> str | None:
    if fs.exists(str(PureWindowsPath(period_root, period))):
        return period
    matches = fs.glob(period_root, period + "*")
    return PureWindowsPath(matches[0][0]).name if len(matches) == 1 else None


def plan_copy(src_root: Path, dst_root: Path, cfg: dict, from_period: str, to_period: str,
              include_skipped: bool):
    """Yield (source, destination, action, size) for everything under the period folder.

    Walks with os.scandir because each round trip to the share is expensive and scandir
    carries the file size along, avoiding a second stat call per file.
    """
    rename_res = [re.compile(p) for p in cfg.get("rename_patterns", [])]
    derived_res = [re.compile(p) for p in cfg.get("derived_patterns", [])]
    skip = {s.lower() for s in cfg.get("skip_contents", [])} if not include_skipped else set()

    def walk(folder: Path, rel: Path, inside_skipped: bool):
        try:
            entries = sorted(os.scandir(folder), key=lambda e: e.name)
        except OSError:
            return
        for entry in entries:
            child_rel = rel / entry.name
            if entry.is_dir(follow_symlinks=False):
                skipped_here = inside_skipped or entry.name.lower() in skip
                yield Path(entry.path), dst_root / child_rel, \
                    "mkdir-empty" if skipped_here else "mkdir", 0
                yield from walk(Path(entry.path), child_rel, skipped_here)
                continue
            if inside_skipped or TEMP_FILE_RE.match(entry.name):
                continue
            name = entry.name
            if any(r.search(name) for r in rename_res):
                name = PERIOD_RE.sub(to_period, name)
            if any(r.search(entry.name) for r in derived_res):
                yield Path(entry.path), dst_root / rel / name, "derive", 0
                continue
            action = "rename" if name != entry.name else "copy"
            yield Path(entry.path), dst_root / rel / name, action, entry.stat().st_size

    yield from walk(src_root, Path(), False)


def execute_copy(items, on_progress=None) -> tuple[int, int]:
    """Create folders and copy files from a plan. Existing files are left alone."""
    copied = skipped = 0
    for src, dst, action, _size in items:
        if action in ("mkdir", "mkdir-empty"):
            dst.mkdir(parents=True, exist_ok=True)
            continue
        if action not in ("copy", "rename"):
            continue
        dst.parent.mkdir(parents=True, exist_ok=True)
        if dst.exists():
            skipped += 1
            continue
        shutil.copy2(src, dst)
        copied += 1
        if on_progress:
            on_progress(dst.name, copied)
    return copied, skipped


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--from", dest="from_period", required=True)
    ap.add_argument("--to", dest="to_period", required=True)
    ap.add_argument("--apply", action="store_true", help="actually write; otherwise dry run")
    ap.add_argument("--include-feedback", action="store_true",
                    help="also copy 业务反馈 contents instead of creating an empty folder")
    ap.add_argument("--config", default=str(ROOT / "config" / "roll.json"))
    args = ap.parse_args()

    cfg = json.loads(Path(args.config).read_text(encoding="utf-8"))
    fs = NetFS(cache_file=str(ROOT / ".cache" / "dirs.json"))
    period_root = cfg["period_root"]

    src_name = resolve_source_folder(period_root, args.from_period, fs)
    if src_name is None:
        sys.exit(f"could not find a single folder for period {args.from_period}")
    src_root = Path(period_root) / src_name
    dst_root = Path(period_root) / args.to_period

    if dst_root.exists() and not args.apply:
        print(f"note: {dst_root} already exists; existing files will be skipped\n")

    items = list(plan_copy(src_root, dst_root, cfg, args.from_period, args.to_period,
                           args.include_feedback))
    files = [(s, d, a, z) for s, d, a, z in items if a in ("copy", "rename")]
    total = sum(z for _, _, _, z in files)

    print(f"source: {src_root}")
    print(f"target: {dst_root}")
    print(f"{len(files)} file(s), {total / 1024 / 1024:.1f} MB\n")
    for s, d, a, _ in items:
        if a == "rename":
            print(f"  RENAME {s.name}\n      -> {d.name}")
    skipped_dirs = [d for _, d, a, _ in items if a == "mkdir-empty"]
    if skipped_dirs:
        print(f"\n{len(skipped_dirs)} folder(s) created empty (contents not copied):")
        for d in skipped_dirs:
            print(f"  {d.relative_to(dst_root)}")

    derived = [d for _, d, a, _ in items if a == "derive"]
    if derived:
        print(f"\n{len(derived)} file(s) NOT copied - cut them from this period's finished "
              f"workbook once it is signed off, then re-point the SCM files:")
        for d in derived:
            print(f"  {d.relative_to(dst_root)}")

    if not args.apply:
        print("\ndry run - nothing written. re-run with --apply to execute.")
        return

    copied, skipped = execute_copy(items)
    fs.invalidate(period_root)
    fs.save()
    print(f"\ncopied {copied} file(s), skipped {skipped} already present")


if __name__ == "__main__":
    main()
