"""Find which sheets reference a given sheet by name, and how often.

Answers "does anything actually read this tab?" without opening Excel.
"""

import argparse
import re
import zipfile
from xml.etree import ElementTree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_ATTR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

FORMULA_RE = re.compile(r"<f[^>]*>(.*?)</f>", re.S)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("--name", required=True, help="sheet name fragment to look for")
    ap.add_argument("--examples", type=int, default=3)
    args = ap.parse_args()

    zf = zipfile.ZipFile(args.workbook)
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = {
        r.get("Id"): r.get("Target")
        for r in ET.fromstring(zf.read("xl/_rels/workbook.xml.rels")).findall(f"{{{NS_PKG_REL}}}Relationship")
    }
    # Matches both 'HFM (2)'!A1 and HFM!A1, but not [1]HFM!A1 which is an external book.
    ref_re = re.compile(r"(?<!\])('[^']*" + re.escape(args.name) + r"[^']*'|\b" + re.escape(args.name) + r"[\w ()]*)!")

    total = 0
    for sh in wb.find(f"{{{NS_MAIN}}}sheets"):
        part = "xl/" + rels[sh.get(f"{{{NS_REL_ATTR}}}id")].lstrip("/").replace("../", "")
        if part not in zf.namelist():
            continue
        xml = zf.read(part).decode("utf-8", "replace")
        hits = [f for f in FORMULA_RE.findall(xml) if ref_re.search(f)]
        if not hits:
            continue
        total += len(hits)
        targets = sorted({m.group(1) for f in hits for m in ref_re.finditer(f)})
        print(f"{sh.get('name')}: {len(hits)} formulas -> {targets}")
        for h in hits[:args.examples]:
            print(f"    {h[:130]}")
    print(f"\ntotal referencing formulas: {total}")


if __name__ == "__main__":
    main()
