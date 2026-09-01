"""Read the first rows of one sheet out of a huge workbook, without downloading it all.

成本收入数据库 is 62 MB packed and 492 MB unpacked, with single sheets over 70 MB, so
anything that loads a whole workbook is unusable on this share. A sheet inside an xlsx
is a zip member that can be decompressed from the front and abandoned once enough rows
have come out, which pulls only the leading chunk across the network.

    python tools/peek_rows.py <workbook>                    # list the sheets and sizes
    python tools/peek_rows.py <workbook> --sheet original --rows 3
"""

from __future__ import annotations

import argparse
import io
import re
import zipfile
from xml.etree import ElementTree as ET

M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
P = "{http://schemas.openxmlformats.org/package/2006/relationships}"

COL_RE = re.compile(r"([A-Z]+)")


def sheet_parts(zf: zipfile.ZipFile) -> list[tuple[str, str, str, int]]:
    """(name, state, package part, uncompressed size) for every sheet."""
    sizes = {i.filename: i.file_size for i in zf.infolist()}
    rels = {r.get("Id"): r.get("Target")
            for r in ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))}
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    out = []
    for s in wb.find(M + "sheets"):
        target = (rels.get(s.get(R + "id")) or "").lstrip("/").replace("../", "")
        part = "xl/" + target
        out.append((s.get("name"), s.get("state") or "visible", part, sizes.get(part, 0)))
    return out


def shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    return ["".join(t.text or "" for t in si.iter(M + "t")) for si in root]


def first_rows(zf: zipfile.ZipFile, part: str, limit: int, strings: list[str]):
    """Stream the sheet and stop as soon as `limit` rows have been parsed."""
    rows = []
    with zf.open(part) as fh:
        # iterparse consumes the stream lazily, so the read stops with the loop.
        for event, el in ET.iterparse(io.BufferedReader(fh, 1 << 16), events=("end",)):
            if el.tag != M + "row":
                continue
            cells = []
            for c in el:
                ref = c.get("r") or ""
                v = c.find(M + "v")
                text = v.text if v is not None else None
                if c.get("t") == "s" and text is not None:
                    idx = int(text)
                    text = strings[idx] if idx < len(strings) else f"<s{idx}>"
                elif c.get("t") == "inlineStr":
                    node = c.find(M + "is")
                    text = "".join(t.text or "" for t in node.iter(M + "t")) if node is not None else None
                if text is not None:
                    cells.append((COL_RE.match(ref).group(1) if COL_RE.match(ref) else ref, text))
            rows.append((el.get("r"), cells))
            el.clear()
            if len(rows) >= limit:
                break
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("--sheet")
    ap.add_argument("--rows", type=int, default=2)
    ap.add_argument("--max-cols", type=int, default=40)
    args = ap.parse_args()

    zf = zipfile.ZipFile(args.workbook)
    sheets = sheet_parts(zf)

    if not args.sheet:
        total = sum(s[3] for s in sheets)
        print(f"{len(sheets)} sheet(s), {total / 1048576:.0f} MB unpacked\n")
        for name, state, _, size in sheets:
            print(f"  {name:<30} {state:<8} {size / 1048576:8.1f} MB")
        return

    match = next((s for s in sheets if s[0].lower() == args.sheet.lower()), None)
    if match is None:
        raise SystemExit(f"no sheet named {args.sheet!r}; have {[s[0] for s in sheets]}")

    strings = shared_strings(zf)
    print(f"{match[0]}  ({match[3] / 1048576:.1f} MB unpacked, "
          f"{len(strings)} shared string(s))\n")
    for ref, cells in first_rows(zf, match[2], args.rows, strings):
        print(f"  row {ref}:")
        for col, text in cells[:args.max_cols]:
            print(f"    {col:>4}  {text[:60]}")
        if len(cells) > args.max_cols:
            print(f"    ... {len(cells) - args.max_cols} more column(s)")
        print()


if __name__ == "__main__":
    main()
