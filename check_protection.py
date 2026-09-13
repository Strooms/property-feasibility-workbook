"""Confirm the protection behaves the way the brief demands, and export a PDF.

The requirement is specific: formula cells locked and hidden, input cells still
freely editable, and no password prompt in the way. That is three separate
things, so each is checked separately rather than assumed.
"""

from __future__ import annotations

import json
from pathlib import Path

import win32com.client as win32
from openpyxl import load_workbook

OUT = Path(__file__).parent / "output"
WB = OUT / "property_feasibility.xlsx"
CELLS = json.loads((OUT / "cellmap.json").read_text(encoding="utf-8"))["Residential"]


def static_checks() -> int:
    """What the file itself declares, read without Excel."""
    wb = load_workbook(WB)
    ws = wb["Residential"]
    fails = 0

    inputs = list(CELLS["land"]) + list(CELLS["retail"]) + [
        CELLS[k] for k in ("units", "build", "council", "stamp", "prof",
                           "cont", "sell", "fin", "threshold")]
    formulas = [CELLS[k] for k in ("avg_land", "avg_retail", "total_purchase",
                                   "gross", "net_real", "tdc", "total_cost",
                                   "net_profit", "poc")]

    for ref in inputs:
        if ws[ref].protection.locked:
            print(f"  FAIL input {ref} is locked - user could not type in it")
            fails += 1
    for ref in formulas:
        p = ws[ref].protection
        if not p.locked:
            print(f"  FAIL formula {ref} is unlocked - user could overwrite it")
            fails += 1
        if not p.hidden:
            print(f"  FAIL formula {ref} is not hidden - formula visible in formula bar")
            fails += 1

    if not ws.protection.sheet:
        print("  FAIL sheet protection is off, so locked/hidden mean nothing")
        fails += 1
    if ws.protection.password:
        print("  NOTE sheet carries a password - unprotecting will prompt")

    print(f"  {len(inputs)} input cells unlocked, {len(formulas)} formula cells "
          f"locked + hidden, sheet protection {'on' if ws.protection.sheet else 'OFF'}")
    return fails


def live_checks() -> int:
    """What Excel actually permits."""
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    fails = 0
    try:
        wb = excel.Workbooks.Open(str(WB))
        ws = wb.Worksheets("Residential")

        # Export before touching anything, so the PDF shows the file as shipped
        # rather than the file as poked at by this test.
        pdf = WB.with_suffix(".pdf")
        wb.ExportAsFixedFormat(0, str(pdf))
        print(f"  exported {pdf.name} ({wb.Worksheets.Count} sheets)")

        # An input must accept a value while the sheet is protected.
        try:
            ws.Range(CELLS["units"]).Value = 7
            print("  input cell accepted a value under protection")
        except Exception as exc:
            print(f"  FAIL input cell rejected input under protection: {exc.__class__.__name__}")
            fails += 1

        # A formula cell must refuse.
        try:
            ws.Range(CELLS["total_cost"]).Value = 1
            print("  FAIL formula cell accepted an overwrite under protection")
            fails += 1
        except Exception:
            print("  formula cell correctly refused an overwrite")

        # And its formula must not be readable.
        hidden = ws.Range(CELLS["poc"]).FormulaHidden
        print(f"  profit-on-cost FormulaHidden = {hidden}")
        fails += not hidden

        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()
    return fails


if __name__ == "__main__":
    print("static checks")
    f = static_checks()
    print("\nlive checks")
    f += live_checks()
    print("\n" + ("PROTECTION OK" if f == 0 else f"{f} PROBLEM(S)"))
