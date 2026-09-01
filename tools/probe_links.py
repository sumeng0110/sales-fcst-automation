"""Inspect an xlsx/xlsm workbook's external links and where they are used.

Reads the OOXML package directly so it runs without third-party dependencies.
"""

import re
import sys
import zipfile
from collections import Counter, defaultdict
from urllib.parse import unquote
from xml.etree import ElementTree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_ATTR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

FUNC_RE = re.compile(r"\b([A-Z][A-Z0-9\.]{1,30})\s*\(")
EXTREF_RE = re.compile(r"\[(\d+)\]")


def parse_rels(zf, part):
    """Return {rId: target} for the given package part."""
    base = part.rsplit("/", 1)
    rels_path = f"{base[0]}/_rels/{base[1]}.rels" if len(base) == 2 else f"_rels/{part}.rels"
    if rels_path not in zf.namelist():
        return {}
    root = ET.fromstring(zf.read(rels_path))
    return {r.get("Id"): r.get("Target") for r in root.findall(f"{{{NS_PKG_REL}}}Relationship")}


def main(path):
    zf = zipfile.ZipFile(path)
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    wb_rels = parse_rels(zf, "xl/workbook.xml")

    sheets = []
    for sh in wb.find(f"{{{NS_MAIN}}}sheets"):
        target = wb_rels.get(sh.get(f"{{{NS_REL_ATTR}}}id"), "")
        part = "xl/" + target.lstrip("/").replace("../", "")
        sheets.append((sh.get("name"), sh.get("state") or "visible", part))

    # Formula tokens like [3]Sheet1!A1 refer to the 3rd entry in externalReferences.
    ext_index_to_file = {}
    ext_refs = wb.find(f"{{{NS_MAIN}}}externalReferences")
    if ext_refs is not None:
        for i, er in enumerate(ext_refs, start=1):
            link_part = "xl/" + wb_rels.get(er.get(f"{{{NS_REL_ATTR}}}id"), "").lstrip("/")
            targets = parse_rels(zf, link_part)
            best = ""
            for t in targets.values():
                t = unquote(t)
                if t.startswith("file:///") or "\\" in t or "/" in t:
                    best = t if len(t) > len(best) else best
            ext_index_to_file[i] = best or next(iter(targets.values()), "?")

    print("=== EXTERNAL REFERENCE INDEX MAP ===")
    for i, f in sorted(ext_index_to_file.items()):
        print(f"[{i}]\t{f}")

    usage = defaultdict(Counter)
    funcs = Counter()
    formula_totals = Counter()
    samples = defaultdict(list)

    print("\n=== PER-SHEET FORMULA / EXTERNAL-LINK USAGE ===")
    for name, state, part in sheets:
        if part not in zf.namelist():
            print(f"{name}\t<missing part {part}>")
            continue
        xml = zf.read(part).decode("utf-8", "replace")
        formulas = re.findall(r"<f[^>]*>(.*?)</f>", xml, re.S)
        formula_totals[name] = len(formulas)
        for f in formulas:
            for fn in FUNC_RE.findall(f):
                funcs[fn] += 1
            for idx in EXTREF_RE.findall(f):
                usage[name][int(idx)] += 1
                if len(samples[(name, int(idx))]) < 2:
                    samples[(name, int(idx))].append(f[:180])
        if formulas:
            refs = ", ".join(f"[{k}]x{v}" for k, v in sorted(usage[name].items()))
            print(f"{name}\t({state})\tformulas={len(formulas)}\t{refs or '-'}")

    print("\n=== TOP FUNCTIONS USED ===")
    for fn, c in funcs.most_common(30):
        print(f"{fn}\t{c}")

    print("\n=== SAMPLE FORMULAS PER (SHEET, LINK) ===")
    for (name, idx), ex in sorted(samples.items()):
        for e in ex:
            print(f"{name}\t[{idx}]\t{e}")

    print(f"\nTOTAL FORMULAS: {sum(formula_totals.values())}")


if __name__ == "__main__":
    main(sys.argv[1])
