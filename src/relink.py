"""Re-point a workbook's external links when rolling forward to a new period.

Excel stores each external link target inside xl/externalLinks/_rels/*.rels, usually
as several equivalent spellings of the same file (absolute UNC, workbook-relative and
server-relative). Rewriting those entries directly avoids launching Excel and lets the
whole plan be previewed before anything is touched.

The rolling rule is deliberately narrow: only files living inside the forecast period
tree (A-P1/CY26/<period>/...) advance with the period. Baselines, prior-year
comparatives, other departments' trees and links to personal drives are pinned, and
anything ambiguous is reported rather than guessed at.
"""

from __future__ import annotations

import calendar
import re
import shutil
import zipfile
from dataclasses import dataclass, field
from pathlib import PureWindowsPath
from urllib.parse import unquote
from xml.etree import ElementTree as ET

NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
ET.register_namespace("", NS_PKG_REL)

# Forecast period tokens such as 6+6, 0+12, 11+1.
PERIOD_RE = re.compile(r"(?<![\d.+])(\d{1,2}\+\d{1,2})(?![\d.+])")

# Month stamps used by the recurring inputs: "Rate 20260630.xlsx", "202606 ... 输入模板.xlsx".
STAMP_DAY_RE = re.compile(r"(?<!\d)(20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])(?!\d)")
STAMP_MONTH_RE = re.compile(r"(?<!\d)(20\d{2})(0[1-9]|1[0-2])(?!\d)")

# Business feedback files are renamed freely every month. Wildcarding the parts
# that churn leaves a stable signature to match this month's equivalent on.
_FUZZY_SUBS = [
    (PERIOD_RE, "*"),
    (STAMP_DAY_RE, "*"),
    (re.compile(r"(?<![A-Za-z\d])(CY\d{2})(0[1-9]|1[0-2])(?!\d)"), r"\1*"),
    (STAMP_MONTH_RE, "*"),
    (re.compile(r"(?<=[ _-])(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])(?!\d)"), "*"),
    (re.compile(r"(?<=[ _-])[Vv]\d+(?![\w])"), "*"),
    # The wo adj copies are named by hand and the separator drifts between periods
    # ("FCST-wo adj" one month, "FCST - wo adj" the next), so it cannot be matched
    # literally.
    (re.compile(r"[-\s]+wo[\s_]*adj", re.I), "*wo*adj"),
]
_STAR_RUN_RE = re.compile(r"\*[\s_-]*\*")


def fuzzy_pattern(stem: str) -> str:
    """Wildcard the volatile parts of a filename so next month's version matches."""
    out = stem
    for rx, rep in _FUZZY_SUBS:
        out = rx.sub(rep, out)
    while _STAR_RUN_RE.search(out):
        out = _STAR_RUN_RE.sub("*", out)
    return out

ROLLED = "rolled"
REPAIRED = "repaired"
PINNED = "pinned"
UNCHANGED = "unchanged"
MISSING = "missing"
WAITING = "waiting"
UNRESOLVED = "unresolved"


@dataclass
class LinkPlan:
    """One decision, covering every spelling of a single external link."""

    link_no: int
    old_abs: str
    new_abs: str
    status: str
    note: str = ""
    rel_ids: list[str] = field(default_factory=list)


def normalize(path: str) -> str:
    """Collapse '.' and '..' segments without touching the filesystem."""
    p = PureWindowsPath(path)
    anchor = p.anchor
    parts: list[str] = []
    for seg in p.parts[1:] if anchor else p.parts:
        if seg == ".":
            continue
        if seg == ".." and parts:
            parts.pop()
        elif seg != "..":
            parts.append(seg)
    return anchor + "\\".join(parts) if anchor else "\\".join(parts)


@dataclass
class Rewriter:
    share_prefix: str
    managed_root: str
    period_root: str
    from_period: str
    to_period: str
    fiscal_year: int = 2026
    months: int = 12
    pinned_periods: set[str] = field(default_factory=set)

    def advance(self, period: str) -> str | None:
        """Next period in the n+(months-n) sequence, or None past year end."""
        n = int(period.split("+")[0])
        return f"{n + 1}+{self.months - n - 1}" if n + 1 <= self.months else None

    @property
    def from_month(self) -> int:
        return int(self.from_period.split("+")[0])

    @property
    def to_month(self) -> int:
        return int(self.to_period.split("+")[0])

    def advance_stamp(self, name: str) -> str:
        """Move a YYYYMMDD or YYYYMM stamp from the source month to the target month.

        Stamps naming any other month belong to a fixed historical input and are left
        alone, so only the recurring monthly files move.
        """
        if self.to_month < 1:
            return name

        def day(m):
            y, mo = int(m.group(1)), int(m.group(2))
            if (y, mo) != (self.fiscal_year, self.from_month):
                return m.group(0)
            last = calendar.monthrange(self.fiscal_year, self.to_month)[1]
            return f"{self.fiscal_year}{self.to_month:02d}{last:02d}"

        def month(m):
            y, mo = int(m.group(1)), int(m.group(2))
            if (y, mo) != (self.fiscal_year, self.from_month):
                return m.group(0)
            return f"{self.fiscal_year}{self.to_month:02d}"

        out = STAMP_DAY_RE.sub(day, name)
        return STAMP_MONTH_RE.sub(month, out)

    def to_absolute(self, target: str, workbook_dir: PureWindowsPath) -> str:
        t = unquote(target).strip()
        if t.lower().startswith("file:///"):
            t = t[len("file:///"):]
        t = t.replace("/", "\\")
        if t.startswith("\\\\"):
            return normalize(t)
        if t.startswith("\\"):
            return normalize(self.share_prefix.rstrip("\\") + t)
        return normalize(str(PureWindowsPath(workbook_dir, t)))

    def is_managed(self, abs_path: str) -> bool:
        return self.managed_root.lower() in abs_path.lower()

    def period_segment(self, abs_path: str) -> str | None:
        """Return the <period> folder name if abs_path sits in the period tree."""
        root = self.period_root.rstrip("\\").lower()
        low = abs_path.lower()
        if not low.startswith(root + "\\"):
            return None
        rest = abs_path[len(root) + 1:]
        seg = rest.split("\\", 1)[0]
        m = PERIOD_RE.match(seg)
        return seg if m else None


def make_folder_resolver(rw: Rewriter, fs):
    """Map a period to its real folder name, which may carry a suffix like '(new)'."""
    cache: dict[str, str | None] = {}

    def resolve(period: str) -> str | None:
        if period not in cache:
            if fs.exists(str(PureWindowsPath(rw.period_root, period))):
                cache[period] = period
            else:
                matches = fs.glob(rw.period_root, period + "*")
                cache[period] = PureWindowsPath(matches[0][0]).name if len(matches) == 1 else None
        return cache[period]

    return resolve


def decide(link_no: int, old_abs: str, rw: Rewriter, fs, folder_for) -> LinkPlan:
    if not rw.is_managed(old_abs):
        return LinkPlan(link_no, old_abs, old_abs, PINNED,
                        "outside managed root (personal drive or other tree)")

    seg = rw.period_segment(old_abs)
    if seg is None:
        # Recurring monthly inputs (FX rate, HFM extracts) live outside the period
        # tree but carry a month stamp that still has to move forward. Only the file
        # name is stamped: dated folder names like "20260509-v6实时更新" are version
        # directories, and inventing a new one would point at nothing.
        p = PureWindowsPath(old_abs)
        stamped = str(PureWindowsPath(p.parent, rw.advance_stamp(p.name)))
        if stamped != old_abs:
            if fs.exists(stamped):
                return LinkPlan(link_no, old_abs, stamped, ROLLED,
                                f"month stamp {rw.from_month:02d} -> {rw.to_month:02d}")
            return LinkPlan(link_no, old_abs, stamped, WAITING,
                            "this month's input file has not been published yet")
        status = PINNED if fs.exists(old_abs) else MISSING
        note = "outside the A-P1 period tree" if status == PINNED else "pinned target no longer exists"
        return LinkPlan(link_no, old_abs, old_abs, status, note)

    seg_period = PERIOD_RE.match(seg).group(1)
    if seg_period in rw.pinned_periods:
        status = UNCHANGED if fs.exists(old_abs) else MISSING
        return LinkPlan(link_no, old_abs, old_abs, status, f"{seg_period} is a pinned baseline")

    next_period = rw.advance(seg_period)
    if next_period is None:
        return LinkPlan(link_no, old_abs, old_abs, PINNED, f"{seg_period} is the last period of the year")

    dest_folder = folder_for(next_period)
    if dest_folder is None:
        return LinkPlan(link_no, old_abs, old_abs, UNRESOLVED,
                        f"no folder found for period {next_period}")

    root_len = len(rw.period_root.rstrip("\\")) + 1
    rest = old_abs[root_len + len(seg):]
    head, _, tail = rest.rpartition("\\")
    stamped_rest = f"{head}\\{rw.advance_stamp(tail)}" if head else rest
    candidate = normalize(str(PureWindowsPath(rw.period_root, dest_folder)) + stamped_rest)
    step = f"{seg_period} -> {next_period}"

    if fs.exists(candidate):
        return LinkPlan(link_no, old_abs, candidate, ROLLED, step)

    # The folder moved on but the file inside was also renamed to the new period.
    p = PureWindowsPath(candidate)
    renamed = str(PureWindowsPath(p.parent, PERIOD_RE.sub(next_period, p.name)))
    if renamed != candidate and fs.exists(renamed):
        return LinkPlan(link_no, old_abs, renamed, REPAIRED,
                        f"{step}, filename advanced to match its folder")

    stem_pattern = fuzzy_pattern(p.stem)
    if "*" in stem_pattern:
        matches = sorted(fs.glob(str(p.parent), stem_pattern + p.suffix), key=lambda x: x[1], reverse=True)
        if len(matches) == 1:
            return LinkPlan(link_no, old_abs, matches[0][0], REPAIRED,
                            f"{step}, matched '{stem_pattern}'")
        if matches:
            names = ", ".join(PureWindowsPath(m).name for m, _ in matches[:4])
            return LinkPlan(link_no, old_abs, matches[0][0], UNRESOLVED,
                            f"{len(matches)} files match '{stem_pattern}': {names}")

    if stamped_rest != rest:
        return LinkPlan(link_no, old_abs, candidate, WAITING,
                        "this month's input file has not been published yet")

    if "*" in stem_pattern and fs.exists(str(p.parent)):
        # Next month's name is unknowable, so keep the old target and say so rather
        # than inventing a path that will silently resolve to nothing.
        return LinkPlan(link_no, old_abs, old_abs, WAITING,
                        f"nothing matches '{stem_pattern}' yet; re-run once this month's file arrives")

    return LinkPlan(link_no, old_abs, candidate, UNRESOLVED, "no candidate file found")


def build_plan(workbook: str, rw: Rewriter, fs, base_dir: str | None = None,
               decide_fn=None) -> list[LinkPlan]:
    """Decide a new target for every external link without writing anything.

    base_dir overrides the folder that workbook-relative links resolve against, so a
    local scratch copy can be analysed as if it sat in its share location.

    decide_fn swaps in a different rule set for workbooks that are not part of the
    forecast chain; see scm.py.
    """
    wb_dir = PureWindowsPath(base_dir) if base_dir else PureWindowsPath(workbook).parent
    folder_for = make_folder_resolver(rw, fs)
    decide_fn = decide_fn or decide
    plans: list[LinkPlan] = []

    with zipfile.ZipFile(workbook) as zf:
        rels_names = sorted(
            n for n in zf.namelist()
            if re.fullmatch(r"xl/externalLinks/_rels/externalLink\d+\.xml\.rels", n)
        )
        for name in rels_names:
            link_no = int(re.search(r"externalLink(\d+)", name).group(1))
            root = ET.fromstring(zf.read(name))
            rels = root.findall(f"{{{NS_PKG_REL}}}Relationship")
            if not rels:
                continue

            # Every spelling names the same workbook, so decide once from the most
            # explicit one and let the rest inherit that decision.
            resolved = [(r.get("Id"), rw.to_absolute(r.get("Target") or "", wb_dir)) for r in rels]
            _, primary = max(resolved, key=lambda x: len(x[1]))
            plan = decide_fn(link_no, primary, rw, fs, folder_for)
            plan.rel_ids = [rid for rid, _ in resolved]
            plans.append(plan)

    return plans


def escape_target(path: str) -> str:
    """Escape a link target the way Excel does.

    Excel percent-encodes spaces but writes non-ASCII characters literally. Using a
    general URL quoter here encodes the Chinese folder names too, and Excel then reads
    the escapes as part of the name, silently breaking every formula behind the link.
    """
    return path.replace("%", "%25").replace(" ", "%20")


def _render_target(raw: str, new_abs: str, wb_dir: PureWindowsPath, share_prefix: str,
                   force_absolute: bool = False) -> str:
    """Rewrite new_abs using the same spelling style the original target used.

    force_absolute drops the relative spellings, which is what you want whenever the
    workbook may be opened from somewhere other than its usual folder.
    """
    if force_absolute:
        return "file:///" + escape_target(new_abs)
    low = raw.lower()
    if low.startswith("file:///"):
        return "file:///" + escape_target(new_abs)
    unq = unquote(raw)
    if unq.startswith("/") and not unq.startswith("//"):
        prefix = share_prefix.rstrip("\\")
        stripped = new_abs[len(prefix):] if new_abs.lower().startswith(prefix.lower()) else new_abs
        return escape_target(stripped.replace("\\", "/"))
    if unq.startswith("\\\\") or (len(unq) > 1 and unq[1] == ":"):
        return new_abs
    base = PureWindowsPath(new_abs)
    try:
        return escape_target(str(base.relative_to(wb_dir)).replace("\\", "/"))
    except ValueError:
        pass
    for depth, anc in enumerate(wb_dir.parents, start=1):
        try:
            rel = base.relative_to(anc)
        except ValueError:
            continue
        return escape_target(("../" * depth) + str(rel).replace("\\", "/"))
    return new_abs


def apply_plan(src: str, dst: str, plans: list[LinkPlan], share_prefix: str,
               base_dir: str | None = None, force_absolute: bool = False) -> int:
    """Write a copy of the workbook with the planned link targets substituted."""
    # WAITING links are rewritten on purpose: the target is where this month's input
    # will land, so Excel picks it up as soon as the file is published.
    changed = {p.link_no: p for p in plans
               if p.status in (ROLLED, REPAIRED, WAITING) and p.new_abs != p.old_abs}
    if not changed:
        shutil.copyfile(src, dst)
        return 0

    wb_dir = PureWindowsPath(base_dir) if base_dir else PureWindowsPath(dst).parent
    written = 0
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            m = re.fullmatch(r"xl/externalLinks/_rels/externalLink(\d+)\.xml\.rels", item.filename)
            if m and int(m.group(1)) in changed:
                plan = changed[int(m.group(1))]
                root = ET.fromstring(data)
                for rel in root.findall(f"{{{NS_PKG_REL}}}Relationship"):
                    rel.set("Target", _render_target(rel.get("Target") or "", plan.new_abs,
                                                     wb_dir, share_prefix, force_absolute))
                    written += 1
                data = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
            zout.writestr(item, data)
    return written
