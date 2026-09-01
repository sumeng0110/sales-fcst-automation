"""Open workbooks in Excel, refresh their external links and recalculate.

The link re-pointing is done by editing the package XML, which is fast but leaves the
cached values untouched. This step is what actually pulls the new numbers through, so
it has to run after the links are correct, and in dependency order.

Requires Excel on the machine (pywin32 talks to it over COM).
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

XL_CALCULATION_AUTOMATIC = -4105
XL_DONE = 0
XL_EXCEL_LINKS = 1
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


def refresh(paths: list[str], visible: bool = False, timeout_s: int = 900) -> list[str]:
    """Open each workbook in turn, update links, recalculate and save."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    saved: list[str] = []
    try:
        app.Visible = visible
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        app.EnableEvents = False
        app.AskToUpdateLinks = False
        # Workbooks in this chain carry macros that must not run during a refresh.
        app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE

        for path in paths:
            print(f"opening {path}", flush=True)
            wb = app.Workbooks.Open(path, UpdateLinks=3, ReadOnly=False)
            try:
                # Excel rejects this property until a workbook is open.
                app.Calculation = XL_CALCULATION_AUTOMATIC
                links = wb.LinkSources(XL_EXCEL_LINKS)
                if links:
                    for link in links:
                        try:
                            wb.UpdateLink(Name=link, Type=XL_EXCEL_LINKS)
                        except Exception as exc:  # a dead link must not abort the run
                            print(f"  link not updated: {link} ({exc})", flush=True)

                app.CalculateFullRebuild()
                deadline = time.time() + timeout_s
                while app.CalculationState != XL_DONE:
                    if time.time() > deadline:
                        raise TimeoutError(f"recalculation exceeded {timeout_s}s")
                    time.sleep(0.5)

                wb.Save()
                saved.append(path)
                print(f"  saved ({len(links or [])} link source(s))", flush=True)
            finally:
                wb.Close(SaveChanges=False)
    finally:
        app.Quit()
        pythoncom.CoUninitialize()
    return saved


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbooks", nargs="+", help="paths in dependency order: P120, P110A, P110")
    ap.add_argument("--visible", action="store_true")
    ap.add_argument("--timeout", type=int, default=900)
    args = ap.parse_args()

    missing = [p for p in args.workbooks if not Path(p).exists()]
    if missing:
        sys.exit("not found:\n  " + "\n  ".join(missing))

    saved = refresh(args.workbooks, visible=args.visible, timeout_s=args.timeout)
    print(f"\nrefreshed {len(saved)} workbook(s)")


if __name__ == "__main__":
    main()
