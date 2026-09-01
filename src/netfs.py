"""Cached filesystem probing for the finance share.

Every lookup on a remote finance share costs real time, and a single roll checks
the same folders dozens of times, so directory listings are cached in memory and
on disk.
"""

from __future__ import annotations

import fnmatch
import json
import os
import time
from pathlib import Path, PureWindowsPath


class NetFS:
    """Directory listings keyed by lower-cased name, holding the name as spelled.

    Matching has to ignore case because links are written in whatever case the author
    typed, but the spelling on disk is what gets written back into a link, so it cannot
    be thrown away.
    """

    def __init__(self, cache_file: str | None = None, ttl_seconds: int = 3600):
        self._dirs: dict[str, dict[str, tuple[str, float]]] = {}
        self._cache_file = Path(cache_file) if cache_file else None
        self._ttl = ttl_seconds
        self._load()

    def _load(self):
        if not self._cache_file or not self._cache_file.exists():
            return
        try:
            blob = json.loads(self._cache_file.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return
        if time.time() - blob.get("saved_at", 0) > self._ttl:
            return
        dirs = blob.get("dirs", {})
        sample = next((v for entries in dirs.values() for v in entries.values()), None)
        if sample is not None and not (isinstance(sample, list) and len(sample) == 2):
            return  # cache written by an older version, without the on-disk spelling
        self._dirs = dirs

    def save(self):
        if not self._cache_file:
            return
        self._cache_file.parent.mkdir(parents=True, exist_ok=True)
        self._cache_file.write_text(
            json.dumps({"saved_at": time.time(), "dirs": self._dirs}, ensure_ascii=False),
            encoding="utf-8",
        )

    def _listing(self, folder: str) -> dict[str, tuple[str, float]]:
        key = folder.rstrip("\\").lower()
        if key in self._dirs:
            return self._dirs[key]
        entries: dict[str, tuple[str, float]] = {}
        try:
            with os.scandir(folder) as it:
                for e in it:
                    try:
                        entries[e.name.lower()] = (e.name, e.stat().st_mtime)
                    except OSError:
                        entries[e.name.lower()] = (e.name, 0.0)
        except (FileNotFoundError, NotADirectoryError, PermissionError, OSError):
            entries = {}
        self._dirs[key] = entries
        return entries

    def exists(self, path: str) -> bool:
        p = PureWindowsPath(path)
        if str(p.parent) in ("", "."):
            return False
        listing = self._listing(str(p.parent))
        if p.name.lower() in listing:
            return True
        # A folder path has no entry of its own until its parent is listed.
        return bool(self._listing(str(p)))

    def glob(self, folder: str, pattern: str) -> list[tuple[str, float]]:
        """Matching files as (path spelled the way it is on disk, mtime)."""
        listing = self._listing(folder)
        pat = pattern.lower()
        return [
            (str(PureWindowsPath(folder, real)), mtime)
            for lower, (real, mtime) in listing.items()
            if fnmatch.fnmatch(lower, pat)
        ]

    def invalidate(self, folder: str):
        self._dirs.pop(folder.rstrip("\\").lower(), None)
