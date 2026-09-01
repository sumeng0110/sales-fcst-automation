"""Local web UI for the monthly forecast roll.

Runs on this machine so it can reach the finance share and drive Excel. Every stage is
split into a plan call and an apply call: the browser shows what would happen, and
nothing is written until you confirm.

    .venv\\Scripts\\python.exe -m uvicorn web.app:app --port 8765
"""

from __future__ import annotations

import shutil
import sys
import threading
import time
import traceback
import uuid
from dataclasses import dataclass, field
from pathlib import Path, PureWindowsPath
from typing import Any, Callable

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "tools"))

import copy_period  # noqa: E402
from hfm_guard import check as hfm_check  # noqa: E402
from netfs import NetFS  # noqa: E402
from relink import (REPAIRED, ROLLED, UNRESOLVED, WAITING, Rewriter, apply_plan,  # noqa: E402
                    build_plan)
from run_month import workbook_paths  # noqa: E402

import json  # noqa: E402

CONFIG_PATH = ROOT / "config" / "roll.json"
app = FastAPI(title="Sales FCST 月度滚动")


def cfg() -> dict:
    return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))


def fresh_fs() -> NetFS:
    """A cache that is dropped after each request; the share changes under us."""
    return NetFS(cache_file=None)


def rewriter(c: dict, from_period: str, to_period: str) -> Rewriter:
    return Rewriter(
        share_prefix=c["share_prefix"],
        managed_root=c["managed_root"],
        period_root=c["period_root"],
        from_period=from_period,
        to_period=to_period,
        fiscal_year=c.get("fiscal_year", 2026),
        months=c.get("months_in_year", 12),
        pinned_periods=set(c.get("pinned_periods", [])),
    )


# ----------------------------------------------------------------- jobs


@dataclass
class Job:
    id: str
    kind: str
    status: str = "running"
    events: list[dict] = field(default_factory=list)
    result: dict = field(default_factory=dict)
    started_at: float = field(default_factory=time.time)
    finished_at: float | None = None


JOBS: dict[str, Job] = {}
JOBS_LOCK = threading.Lock()


def start_job(kind: str, work: Callable[[Callable[[str, str], None]], dict]) -> Job:
    job = Job(id=uuid.uuid4().hex[:12], kind=kind)
    with JOBS_LOCK:
        JOBS[job.id] = job

    def log(level: str, message: str):
        job.events.append({"t": time.time(), "level": level, "message": message})

    def runner():
        try:
            job.result = work(log) or {}
            job.status = "done"
        except Exception as exc:
            log("error", f"{type(exc).__name__}: {exc}")
            log("error", traceback.format_exc(limit=4))
            job.status = "error"
        finally:
            job.finished_at = time.time()

    threading.Thread(target=runner, daemon=True).start()
    return job


@app.get("/api/jobs/{job_id}")
def job_status(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "unknown job")
    return {
        "id": job.id,
        "kind": job.kind,
        "status": job.status,
        "events": job.events,
        "result": job.result,
        "elapsed": round((job.finished_at or time.time()) - job.started_at, 1),
    }


# ----------------------------------------------------------------- periods


@app.get("/api/periods")
def list_periods():
    c = cfg()
    root = Path(c["period_root"])
    if not root.exists():
        raise HTTPException(503, f"share not reachable: {root}")
    out = []
    for d in root.iterdir():
        if not d.is_dir():
            continue
        token = d.name.split("(")[0].strip()
        if "+" not in token:
            continue
        try:
            index = int(token.split("+")[0])
        except ValueError:
            continue
        out.append({
            "folder": d.name,
            "period": token,
            "index": index,
            "modified": d.stat().st_mtime,
        })
    out.sort(key=lambda x: x["index"])
    months = c.get("months_in_year", 12)
    latest = out[-1] if out else None
    return {
        "periods": out,
        "suggested_from": latest["period"] if latest else None,
        "suggested_to": (f"{latest['index'] + 1}+{months - latest['index'] - 1}"
                         if latest and latest["index"] < months else None),
        "fiscal_year": c.get("fiscal_year"),
        "period_root": c["period_root"],
    }


# ----------------------------------------------------------------- stage 1: copy


class PeriodPair(BaseModel):
    from_period: str
    to_period: str


def _copy_items(c: dict, body: PeriodPair, include_feedback: bool):
    fs = fresh_fs()
    src_name = copy_period.resolve_source_folder(c["period_root"], body.from_period, fs)
    if src_name is None:
        raise HTTPException(400, f"no single folder found for period {body.from_period}")
    src_root = Path(c["period_root"]) / src_name
    dst_root = Path(c["period_root"]) / body.to_period
    items = list(copy_period.plan_copy(src_root, dst_root, c, body.from_period,
                                       body.to_period, include_feedback))
    return src_root, dst_root, items


@app.post("/api/copy/plan")
def copy_plan(body: PeriodPair, include_feedback: bool = False):
    c = cfg()
    src_root, dst_root, items = _copy_items(c, body, include_feedback)
    files = [(s, d, a, z) for s, d, a, z in items if a in ("copy", "rename")]
    return {
        "source": str(src_root),
        "target": str(dst_root),
        "target_exists": dst_root.exists(),
        "file_count": len(files),
        "total_mb": round(sum(z for _, _, _, z in files) / 1048576, 1),
        "renames": [{"from": s.name, "to": d.name} for s, d, a, _ in items if a == "rename"],
        "empty_folders": [str(d.relative_to(dst_root)) for _, d, a, _ in items
                          if a == "mkdir-empty"],
        "derived": [str(d.relative_to(dst_root)) for _, d, a, _ in items if a == "derive"],
        "files": [{"name": d.name, "mb": round(z / 1048576, 2),
                   "relative": str(d.relative_to(dst_root))} for s, d, a, z in files],
    }


@app.post("/api/copy/apply")
def copy_apply(body: PeriodPair, include_feedback: bool = False):
    c = cfg()
    _, dst_root, items = _copy_items(c, body, include_feedback)
    total = sum(1 for _, _, a, _ in items if a in ("copy", "rename"))

    def work(log):
        log("info", f"copying {total} file(s) to {dst_root}")
        copied, skipped = execute_with_log(items, log, total)
        log("info", f"copied {copied}, skipped {skipped} already present")
        return {"copied": copied, "skipped": skipped, "target": str(dst_root)}

    return {"job": start_job("copy", work).id}


def execute_with_log(items, log, total):
    def on_progress(name, n):
        if n % 3 == 0 or n == total:
            log("info", f"  {n}/{total}  {name}")

    return copy_period.execute_copy(items, on_progress=on_progress)


# ----------------------------------------------------------------- stage 2: links


@app.post("/api/links/plan")
def links_plan(body: PeriodPair):
    c = cfg()
    rw = rewriter(c, body.from_period, body.to_period)
    fs = fresh_fs()
    out = []
    for wb_id, path in workbook_paths(c, body.to_period):
        if not path.exists():
            out.append({"workbook": wb_id, "file": path.name, "missing": True, "links": []})
            continue
        plans = build_plan(str(path), rw, fs, base_dir=str(path.parent))
        out.append({
            "workbook": wb_id,
            "file": path.name,
            "path": str(path),
            "missing": False,
            "links": [{
                "no": p.link_no,
                "status": p.status,
                "note": p.note,
                "old": p.old_abs,
                "new": p.new_abs,
                "old_name": PureWindowsPath(p.old_abs).name,
                "new_name": PureWindowsPath(p.new_abs).name,
                "changed": p.new_abs != p.old_abs,
            } for p in sorted(plans, key=lambda x: x.link_no)],
        })
    blocking = sum(1 for wb in out for l in wb["links"] if l["status"] == UNRESOLVED)
    waiting = sum(1 for wb in out for l in wb["links"] if l["status"] == WAITING)
    return {"workbooks": out, "unresolved": blocking, "waiting": waiting}


@app.post("/api/links/apply")
def links_apply(body: PeriodPair):
    c = cfg()
    rw = rewriter(c, body.from_period, body.to_period)
    targets = workbook_paths(c, body.to_period)

    def work(log):
        fs = fresh_fs()
        written = {}
        for wb_id, path in targets:
            if not path.exists():
                raise FileNotFoundError(f"{path} - run the copy stage first")
            log("info", f"{wb_id}: {path.name}")
            plans = build_plan(str(path), rw, fs, base_dir=str(path.parent))
            unresolved = [p for p in plans if p.status == UNRESOLVED]
            if unresolved:
                for p in unresolved:
                    log("error", f"  unresolved [{p.link_no}] {p.note}")
                raise RuntimeError(f"{wb_id} has {len(unresolved)} unresolved link(s)")
            changed = [p for p in plans if p.status in (ROLLED, REPAIRED, WAITING)
                       and p.new_abs != p.old_abs]
            for p in changed:
                log("info", f"  [{p.link_no}] -> {PureWindowsPath(p.new_abs).name}")
            if changed:
                tmp = path.with_suffix(path.suffix + ".tmp")
                apply_plan(str(path), str(tmp), plans, c["share_prefix"], base_dir=str(path.parent))
                tmp.replace(path)
            written[wb_id] = len(changed)
            log("info", f"  {len(changed)} link(s) rewritten")
        return {"written": written}

    return {"job": start_job("links", work).id}


# ----------------------------------------------------------------- stage 3: HFM


@app.get("/api/hfm/check")
def hfm_status(to_period: str):
    c = cfg()
    actual_months = int(to_period.split("+")[0])
    out = []
    for wb_id, path in workbook_paths(c, to_period):
        if not path.exists():
            out.append({"workbook": wb_id, "missing": True, "grids": []})
            continue
        grids = []
        for sheet, cols, verdict in hfm_check(str(path), c.get("fiscal_year", 2026), actual_months):
            if verdict == "not in use":
                continue
            grids.append({
                "sheet": sheet,
                "verdict": verdict,
                "ok": verdict.startswith("ok"),
                "months": [{"period": mc.period, "year": mc.year, "populated": mc.populated,
                            "invalid": mc.invalid}
                           for mc in (cols or []) if mc.year == c.get("fiscal_year", 2026)],
            })
        out.append({"workbook": wb_id, "missing": False, "grids": grids})
    stale = [f"{wb['workbook']} / {g['sheet']}" for wb in out for g in wb["grids"] if not g["ok"]]
    return {"workbooks": out, "stale": stale, "expected_month": actual_months}


# ----------------------------------------------------------------- stage 4: recalc


@app.post("/api/recalc")
def recalc(body: PeriodPair, skip_hfm_check: bool = False):
    c = cfg()
    targets = workbook_paths(c, body.to_period)
    actual_months = int(body.to_period.split("+")[0])

    def work(log):
        if not skip_hfm_check:
            stale = []
            for wb_id, path in targets:
                for sheet, _, verdict in hfm_check(str(path), c.get("fiscal_year", 2026), actual_months):
                    if verdict != "not in use" and not verdict.startswith("ok"):
                        stale.append(f"{wb_id} / {sheet}: {verdict}")
            if stale:
                for s in stale:
                    log("error", s)
                raise RuntimeError("Smart View grids still hold an earlier month")
        from excel_refresh import refresh
        log("info", "starting Excel; this takes 10-25 minutes for the full chain")
        paths = [str(p) for _, p in targets]
        saved = refresh(paths)
        log("info", f"recalculated and saved {len(saved)} workbook(s)")
        return {"saved": saved}

    return {"job": start_job("recalc", work).id}


# ----------------------------------------------------------------- stage 5: compare


class ComparePair(BaseModel):
    left: str
    right: str
    tolerance: float = 0.5


@app.post("/api/compare")
def compare(body: ComparePair):
    from compare_cache import live_values, norm

    def work(log):
        log("info", "reading both workbooks")
        a, b = live_values(body.left), live_values(body.right)
        sheets = sorted(set(a) & set(b))
        rows, total_same, total_diff = [], 0, 0
        for sheet in sheets:
            same = diff = 0
            examples = []
            for ref in set(a[sheet]) | set(b[sheet]):
                va, vb = a[sheet].get(ref), b[sheet].get(ref)
                if va is None or vb is None:
                    continue
                ta, na = norm(*va)
                tb, nb = norm(*vb)
                ok = (abs(na - nb) <= body.tolerance if ta == "num" and tb == "num"
                      else str(na) == str(nb))
                if ok:
                    same += 1
                else:
                    diff += 1
                    if len(examples) < 3:
                        examples.append({"cell": ref, "left": str(na), "right": str(nb)})
            total_same += same
            total_diff += diff
            rows.append({"sheet": sheet, "same": same, "diff": diff, "examples": examples})
        log("info", f"compared {total_same + total_diff} cells")
        return {
            "rows": sorted(rows, key=lambda r: -r["diff"]),
            "same": total_same,
            "diff": total_diff,
            "only_left": sorted(set(a) - set(b)),
            "only_right": sorted(set(b) - set(a)),
        }

    return {"job": start_job("compare", work).id}


# ----------------------------------------------------------------- uploads


@app.post("/api/upload")
async def upload(to_period: str = Form(...), subfolder: str = Form(...),
                 file: UploadFile = File(...)):
    """Drop a file that arrived by email into the right folder of the new period."""
    c = cfg()
    base = Path(c["period_root"]) / to_period
    if not base.exists():
        raise HTTPException(400, f"{base} does not exist yet - run the copy stage first")
    target_dir = (base / subfolder).resolve()
    if not str(target_dir).lower().startswith(str(base.resolve()).lower()):
        raise HTTPException(400, "subfolder must stay inside the period folder")
    target_dir.mkdir(parents=True, exist_ok=True)
    dest = target_dir / Path(file.filename).name
    if dest.exists():
        raise HTTPException(409, f"{dest.name} already exists")
    with dest.open("wb") as fh:
        shutil.copyfileobj(file.file, fh)
    return {"saved": str(dest), "bytes": dest.stat().st_size}


# ----------------------------------------------------------------- static


STATIC = Path(__file__).resolve().parent / "static"
app.mount("/static", StaticFiles(directory=STATIC), name="static")


@app.get("/")
def index():
    return FileResponse(STATIC / "index.html")
