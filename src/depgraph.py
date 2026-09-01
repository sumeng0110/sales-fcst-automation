"""Build the dependency closure of a workbook by following external links backwards.

Starts from the final deliverable and walks link by link, so only files that actually
feed the result are ever opened. Nothing else on the share is touched, and each
workbook is read once - the OOXML parts holding the link targets are a few kilobytes,
so a node costs one small read even when the file itself is megabytes.

Files that cannot be read (personal drives, deleted sources) become leaves and are
reported rather than skipped silently: an unreadable input is exactly the kind of
thing that quietly poisons a forecast.
"""

from __future__ import annotations

import argparse
import io
import json
import sys
import time
import zipfile
from collections import deque
from dataclasses import dataclass, field
from pathlib import Path, PureWindowsPath
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parent))

from relink import PERIOD_RE, normalize  # noqa: E402

NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
READABLE = {".xlsx", ".xlsm", ".xlsb", ".xls"}


@dataclass
class Node:
    path: str
    name: str
    depth: int
    exists: bool
    size: int = 0
    readable: bool = True
    error: str = ""
    period: str | None = None
    children: list[str] = field(default_factory=list)


def link_targets(path: str, share_prefix: str) -> list[str]:
    """Absolute targets of every external link, read straight from the package.

    The file is pulled into memory in one sequential read before being unzipped.
    Handing the network path to zipfile directly makes it seek per entry, and every
    seek is an SMB round trip - that turns a two second read into several minutes.
    """
    out: list[str] = []
    wb_dir = PureWindowsPath(path).parent
    blob = Path(path).read_bytes()
    with zipfile.ZipFile(io.BytesIO(blob)) as zf:
        names = [n for n in zf.namelist()
                 if n.startswith("xl/externalLinks/_rels/") and n.endswith(".rels")]
        for name in names:
            rels = ET.fromstring(zf.read(name)).findall(f"{{{NS_PKG_REL}}}Relationship")
            spellings = []
            for rel in rels:
                t = (rel.get("Target") or "").strip()
                if not t:
                    continue
                from urllib.parse import unquote
                t = unquote(t)
                if t.lower().startswith("file:///"):
                    t = t[len("file:///"):]
                t = t.replace("/", "\\")
                if t.startswith("\\\\"):
                    spellings.append(normalize(t))
                elif t.startswith("\\"):
                    spellings.append(normalize(share_prefix.rstrip("\\") + t))
                else:
                    spellings.append(normalize(str(PureWindowsPath(wb_dir, t))))
            if spellings:
                # Every spelling names the same workbook; the longest is the explicit one.
                out.append(max(spellings, key=len))
    return out


def period_of(path: str, period_root: str) -> str | None:
    root = period_root.rstrip("\\").lower()
    if not path.lower().startswith(root + "\\"):
        return None
    seg = path[len(root) + 1:].split("\\", 1)[0]
    m = PERIOD_RE.match(seg)
    return m.group(1) if m else None


def is_terminal_input(path: str, leaf_folders: list[str]) -> bool:
    """Files a colleague sent us are inputs, not part of the production chain.

    A feedback workbook is usually a whole copy of the model with its own 20+ links,
    so recursing into it explodes the graph into hundreds of files while telling us
    nothing: we consume its numbers, we do not rebuild it.
    """
    low = path.lower()
    return any(f"\\{folder.lower()}\\" in low for folder in leaf_folders)


def build(root_workbook: str, share_prefix: str, period_root: str, max_depth: int = 3,
          leaf_folders: list[str] | None = None, on_visit=None) -> dict[str, Node]:
    leaf_folders = leaf_folders if leaf_folders is not None else ["业务反馈"]
    nodes: dict[str, Node] = {}
    queue = deque([(normalize(root_workbook), 0)])

    while queue:
        path, depth = queue.popleft()
        key = path.lower()
        if key in nodes:
            nodes[key].depth = min(nodes[key].depth, depth)
            continue

        p = Path(path)
        node = Node(path=path, name=p.name, depth=depth, exists=False,
                    period=period_of(path, period_root))
        nodes[key] = node

        try:
            node.exists = p.exists()
        except OSError as exc:
            node.error = f"unreachable: {exc}"
        if node.exists:
            node.size = p.stat().st_size

        if on_visit:
            on_visit(node, len(nodes), len(queue))

        if not node.exists:
            node.error = node.error or "file not found"
            continue
        if p.suffix.lower() not in READABLE:
            node.readable = False
            continue
        if is_terminal_input(path, leaf_folders):
            node.readable = False
            node.error = "terminal input (business feedback)"
            continue
        if depth >= max_depth:
            node.readable = False
            node.error = f"stopped at depth {max_depth}"
            continue

        try:
            targets = link_targets(path, share_prefix)
        except (zipfile.BadZipFile, OSError, ET.ParseError) as exc:
            node.readable = False
            node.error = f"cannot read links: {type(exc).__name__}"
            continue

        for t in targets:
            node.children.append(t)
            queue.append((t, depth + 1))

    return nodes


def print_tree(nodes: dict[str, Node], root: str, seen: set[str] | None = None,
               prefix: str = "", is_last: bool = True):
    seen = seen if seen is not None else set()
    node = nodes.get(root.lower())
    if node is None:
        return
    mark = "└─ " if is_last else "├─ "
    if not node.exists:
        state = "  [MISSING]"
    elif not node.readable:
        state = f"  [{node.error or 'leaf'}]"
    else:
        state = f"  ({node.size / 1048576:.1f} MB)"
    repeat = " ↩" if root.lower() in seen else ""
    print(f"{prefix}{mark if prefix else ''}{node.name}{state}{repeat}")
    if repeat:
        return
    seen.add(root.lower())
    child_prefix = prefix + ("   " if is_last else "│  ") if prefix else "  "
    for i, child in enumerate(node.children):
        print_tree(nodes, child, seen, child_prefix, i == len(node.children) - 1)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook", help="the final deliverable to trace back from")
    ap.add_argument("--config", default=str(Path(__file__).resolve().parent.parent
                                            / "config" / "roll.json"))
    ap.add_argument("--max-depth", type=int, default=3)
    ap.add_argument("--leaf-folder", action="append",
                    help="folders whose files are inputs, not to be traced into")
    ap.add_argument("--json", dest="json_out")
    args = ap.parse_args()

    cfg = json.loads(Path(args.config).read_text(encoding="utf-8"))

    started = time.time()

    def progress(node, visited, pending):
        state = "ok" if node.exists else "MISSING"
        size = f"{node.size / 1048576:.1f}MB" if node.size else ""
        print(f"[{time.time() - started:6.0f}s | {visited:>3} seen, {pending:>3} queued] "
              f"d{node.depth} {state:<7} {size:>7}  {node.name}", flush=True)

    nodes = build(args.workbook, cfg["share_prefix"], cfg["period_root"], args.max_depth,
                  leaf_folders=args.leaf_folder or cfg.get("leaf_folders", ["业务反馈"]),
                  on_visit=progress)
    print()

    print("=== DEPENDENCY TREE ===")
    print_tree(nodes, normalize(args.workbook))

    readable = [n for n in nodes.values() if n.exists and n.readable]
    missing = [n for n in nodes.values() if not n.exists]
    leaves = [n for n in nodes.values() if n.exists and not n.readable]

    print(f"\n=== SUMMARY ===")
    print(f"files in closure : {len(nodes)}")
    print(f"readable         : {len(readable)}  ({sum(n.size for n in readable) / 1048576:.0f} MB)")
    print(f"leaves           : {len(leaves)}")
    print(f"missing          : {len(missing)}")
    print(f"max depth reached: {max(n.depth for n in nodes.values())}")

    if missing:
        print("\n=== MISSING (links that resolve to nothing) ===")
        for n in sorted(missing, key=lambda x: x.depth):
            print(f"  d{n.depth}  {n.path}")

    by_period: dict[str | None, int] = {}
    for n in nodes.values():
        by_period[n.period] = by_period.get(n.period, 0) + 1
    print("\n=== BY PERIOD FOLDER ===")
    for period, count in sorted(by_period.items(), key=lambda kv: (kv[0] is None, kv[0] or "")):
        print(f"  {period or 'outside period tree'}: {count}")

    if args.json_out:
        Path(args.json_out).write_text(json.dumps(
            {k: v.__dict__ for k, v in nodes.items()}, ensure_ascii=False, indent=2),
            encoding="utf-8")
        print(f"\nwrote {args.json_out}")


if __name__ == "__main__":
    main()
