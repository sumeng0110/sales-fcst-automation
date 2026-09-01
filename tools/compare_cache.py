"""Compare an external link's cached snapshot against the live source workbook.

Excel stores the last-seen values of a linked workbook inside xl/externalLinks/.
When the link target goes missing, those stale values are what the formulas keep
returning, so diffing them against the real file reveals silently wrong numbers.
"""

import sys
import zipfile
from xml.etree import ElementTree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_REL_ATTR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Rounding tolerance: these are financial figures carried at full float precision.
TOL = 0.5


def cached_values(zf, link_no):
    """Return {sheet_name: {cell_ref: value}} recorded for one external link."""
    root = ET.fromstring(zf.read(f"xl/externalLinks/externalLink{link_no}.xml"))
    book = root.find(f"{{{NS_MAIN}}}externalBook")
    names = [s.get("val") for s in book.find(f"{{{NS_MAIN}}}sheetNames")]
    out = {}
    data_set = book.find(f"{{{NS_MAIN}}}sheetDataSet")
    if data_set is None:
        return out
    for sd in data_set:
        sheet = names[int(sd.get("sheetId"))]
        cells = {}
        for row in sd:
            for c in row:
                v = c.find(f"{{{NS_MAIN}}}v")
                if v is None or v.text is None:
                    continue
                cells[c.get("r")] = (c.get("t") or "n", v.text)
        out[sheet] = cells
    return out


def live_values(path):
    """Return {sheet_name: {cell_ref: value}} as currently stored in a workbook."""
    zf = zipfile.ZipFile(path)
    shared = []
    if "xl/sharedStrings.xml" in zf.namelist():
        ss = ET.fromstring(zf.read("xl/sharedStrings.xml"))
        for si in ss:
            shared.append("".join(t.text or "" for t in si.iter(f"{{{NS_MAIN}}}t")))

    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = {
        r.get("Id"): r.get("Target")
        for r in ET.fromstring(zf.read("xl/_rels/workbook.xml.rels")).findall(f"{{{NS_PKG_REL}}}Relationship")
    }
    out = {}
    for sh in wb.find(f"{{{NS_MAIN}}}sheets"):
        part = "xl/" + rels[sh.get(f"{{{NS_REL_ATTR}}}id")].lstrip("/").replace("../", "")
        if part not in zf.namelist():
            continue
        cells = {}
        for _, c in ET.iterparse(zf.open(part)):
            if c.tag != f"{{{NS_MAIN}}}c":
                continue
            v = c.find(f"{{{NS_MAIN}}}v")
            if v is None or v.text is None:
                continue
            t = c.get("t") or "n"
            val = shared[int(v.text)] if t == "s" and v.text.isdigit() else v.text
            cells[c.get("r")] = (t, val)
        out[sh.get("name")] = cells
    return out


def norm(t, v):
    if t in ("s", "str", "inlineStr", "e"):
        return ("text", str(v).strip())
    try:
        return ("num", float(v))
    except ValueError:
        return ("text", str(v).strip())


def main(target_wb, link_no, source_wb):
    cache = cached_values(zipfile.ZipFile(target_wb), link_no)
    live = live_values(source_wb)

    print(f"cached sheets : {sorted(cache)}")
    print(f"live sheets   : {sorted(live)}\n")

    grand_same = grand_diff = grand_missing = 0
    for sheet, cells in sorted(cache.items()):
        if sheet not in live:
            print(f"{sheet}: SHEET NOT IN LIVE FILE (cached cells: {len(cells)})")
            grand_missing += len(cells)
            continue
        same = diff = 0
        examples = []
        for ref, (t, v) in cells.items():
            ct, cv = norm(t, v)
            if ref not in live[sheet]:
                diff += 1
                if len(examples) < 4:
                    examples.append(f"    {ref}: cached={cv!r} live=<empty>")
                continue
            lt, lv = norm(*live[sheet][ref])
            if ct == "num" and lt == "num":
                ok = abs(cv - lv) <= TOL
            else:
                ok = str(cv) == str(lv)
            if ok:
                same += 1
            else:
                diff += 1
                if len(examples) < 4:
                    examples.append(f"    {ref}: cached={cv!r} live={lv!r}")
        grand_same += same
        grand_diff += diff
        flag = "OK" if diff == 0 else "MISMATCH"
        print(f"{sheet}: {flag}  same={same} diff={diff}")
        for e in examples:
            print(e)

    total = grand_same + grand_diff + grand_missing
    print(f"\nTOTAL cached cells={total} same={grand_same} diff={grand_diff} missing_sheet={grand_missing}")
    if total:
        print(f"MATCH RATE: {grand_same / total:.1%}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3])
