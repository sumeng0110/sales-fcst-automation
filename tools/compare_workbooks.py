"""Diff two workbooks cell by cell using their stored values.

Used to check an automated result against the hand-made one: same sheets, same cells,
same numbers. Reports per sheet so a single broken area is obvious.
"""

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from compare_cache import live_values, norm


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("left")
    ap.add_argument("right")
    ap.add_argument("--tolerance", type=float, default=0.5)
    ap.add_argument("--examples", type=int, default=3)
    ap.add_argument("--sheet", action="append", help="limit to these sheets")
    args = ap.parse_args()

    a = live_values(args.left)
    b = live_values(args.right)

    only_a = sorted(set(a) - set(b))
    only_b = sorted(set(b) - set(a))
    if only_a:
        print(f"sheets only in left : {only_a}")
    if only_b:
        print(f"sheets only in right: {only_b}")

    sheets = sorted(set(a) & set(b))
    if args.sheet:
        wanted = {s.lower() for s in args.sheet}
        sheets = [s for s in sheets if s.lower() in wanted]

    total_same = total_diff = 0
    worst: list[tuple[float, str, str, object, object]] = []

    print(f"\n{'sheet':<34} {'same':>7} {'diff':>7}")
    for sheet in sheets:
        same = diff = 0
        examples = []
        for ref in set(a[sheet]) | set(b[sheet]):
            va, vb = a[sheet].get(ref), b[sheet].get(ref)
            if va is None or vb is None:
                if (va or vb) not in (None, ("n", "0"), ("n", "")):
                    diff += 1
                    if len(examples) < args.examples:
                        examples.append(f"    {ref}: left={va} right={vb}")
                continue
            ta, na = norm(*va)
            tb, nb = norm(*vb)
            if ta == "num" and tb == "num":
                delta = abs(na - nb)
                ok = delta <= args.tolerance
                if not ok:
                    worst.append((delta, sheet, ref, na, nb))
            else:
                ok = str(na) == str(nb)
            if ok:
                same += 1
            else:
                diff += 1
                if len(examples) < args.examples:
                    examples.append(f"    {ref}: left={na!r} right={nb!r}")
        total_same += same
        total_diff += diff
        marker = "" if diff == 0 else "  <-- differs"
        print(f"{sheet:<34} {same:>7} {diff:>7}{marker}")
        for e in examples:
            print(e)

    total = total_same + total_diff
    print(f"\ntotal cells compared: {total}   same={total_same}   diff={total_diff}")
    if total:
        print(f"match rate: {total_same / total:.2%}")
    if worst:
        print("\nlargest numeric gaps:")
        for delta, sheet, ref, na, nb in sorted(worst, reverse=True)[:10]:
            print(f"  {sheet}!{ref}: {na:,.1f} vs {nb:,.1f}  (Δ {delta:,.1f})")

    sys.exit(1 if total_diff else 0)


if __name__ == "__main__":
    main()
