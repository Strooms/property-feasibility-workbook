"""Pressure-test the workbook: drive Excel, feed it deals, compare to an
independent Python model.

The brief asks for totals and a verdict that are correct at the edges, not
merely error-free. So the reference numbers here are computed separately, in
Python, rather than read back out of the same formulas being tested.

    python verify.py
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import win32com.client as win32

OUT = Path(__file__).parent / "output"
WB = OUT / "property_feasibility.xlsx"
CELLMAP = OUT / "cellmap.json"

# The workbook may carry a sheet-protection password. Unprotect() without one
# raises a modal password dialog that hangs Excel with no window to dismiss, so
# the password is always passed explicitly. Override with --password.
DEFAULT_PASSWORD = "feas2026"


def unprotect(ws, password: str) -> None:
    try:
        ws.Unprotect(password)
    except Exception:
        # Unprotected sheets reject a password argument; that is fine.
        pass


def expected(d: dict) -> dict:
    """Independent reference model."""
    avg_land = sum(d["land"]) / len(d["land"])
    avg_retail = sum(d["retail"]) / len(d["retail"])
    land_purchase = avg_land * d["units"]
    total_purchase = land_purchase + land_purchase * d["stamp"]

    gross = avg_retail * d["units"]
    net_real = gross - gross * d["sell"]

    build_total = d["build"] * d["units"]
    council_total = d["council"] * d["units"]
    prof = build_total * d["prof"]
    cont = build_total * d["cont"]
    fin = (total_purchase + build_total + council_total + prof + cont) * d["fin"]
    tdc = build_total + council_total + prof + cont + fin

    total_cost = total_purchase + tdc
    net_profit = net_real - total_cost
    poc = net_profit / total_cost if total_cost else 0.0
    verdict = "MEETS THRESHOLD" if poc >= d["threshold"] else "BELOW THRESHOLD"
    return {"total_purchase": total_purchase, "gross": gross, "net_real": net_real,
            "tdc": tdc, "total_cost": total_cost, "net_profit": net_profit,
            "poc": poc, "verdict": verdict}


# Three deals, mirroring what the brief says the client will supply.
BASE = {
    "units": 6, "build": 310000, "council": 28000,
    "stamp": 0.055, "prof": 0.07, "cont": 0.05, "sell": 0.025,
    "fin": 0.06, "threshold": 0.25,
}

# Retail value per dwelling that lands profit-on-cost exactly on 25%, given
# land at 300,000 and the BASE percentages. Derived, not guessed:
#   cost/unit = 1.06 * (1.055*L + 1.12*B + C)
#   poc = 0.975*R / (cost/unit) - 1   ->   R at poc=0.25
_COST_PER_UNIT = 1.06 * (1.055 * 300000 + 1.12 * BASE["build"] + BASE["council"])
R_ON_THRESHOLD = 1.25 * _COST_PER_UNIT / 0.975


def deal(land, retail, **over):
    d = dict(BASE)
    d.update(over)
    d["land"] = [land] * 4
    d["retail"] = [retail] * 4
    return d


DEALS = {
    "healthy deal": deal(300000, 965000),
    "losing deal": deal(700000, 880000, units=5, build=420000, council=45000,
                        prof=0.09, cont=0.08, sell=0.03, fin=0.075),
    "a hair above threshold": deal(300000, R_ON_THRESHOLD + 50),
    "a hair below threshold": deal(300000, R_ON_THRESHOLD - 50),
}


def write_deal(ws, cells: dict, d: dict) -> None:
    for cell, value in zip(cells["land"], d["land"]):
        ws.Range(cell).Value = value
    for cell, value in zip(cells["retail"], d["retail"]):
        ws.Range(cell).Value = value
    for key in ("units", "build", "council", "stamp", "prof", "cont",
                "sell", "fin", "threshold"):
        ws.Range(cells[key]).Value = d[key]


def main() -> int:
    if not WB.exists() or not CELLMAP.exists():
        print(f"Workbook or cell map missing under {OUT}  (run build_workbook.py first)")
        return 1

    CELLS = json.loads(CELLMAP.read_text(encoding="utf-8"))["Residential"]

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    failures = 0

    try:
        wb = excel.Workbooks.Open(str(WB))
        ws = wb.Worksheets("Residential")
        unprotect(ws, DEFAULT_PASSWORD)

        for name, deal in DEALS.items():
            write_deal(ws, CELLS, deal)
            excel.CalculateFull()
            exp = expected(deal)

            print(f"\n{name}")
            print(f"{'':24s}{'excel':>16s}{'expected':>16s}")
            for key in ("total_purchase", "gross", "net_real", "tdc",
                        "total_cost", "net_profit"):
                got = float(ws.Range(CELLS[key]).Value)
                want = exp[key]
                ok = abs(got - want) < 0.51
                failures += not ok
                flag = "" if ok else "   <-- MISMATCH"
                print(f"  {key:22s}{got:16,.0f}{want:16,.0f}{flag}")

            got_poc = float(ws.Range(CELLS["poc"]).Value)
            ok = abs(got_poc - exp["poc"]) < 1e-9
            failures += not ok
            print(f"  {'profit on cost':22s}{got_poc:15.2%}{exp['poc']:16.2%}"
                  f"{'' if ok else '   <-- MISMATCH'}")

            got_verdict = str(ws.Range(CELLS["verdict"]).Value)
            ok = got_verdict == exp["verdict"]
            failures += not ok
            print(f"  {'verdict':22s}{got_verdict:>16s}{exp['verdict']:>16s}"
                  f"{'' if ok else '   <-- MISMATCH'}")

        # Edge cases the brief does not mention but a public lead magnet will meet.
        print("\nedge cases")
        ws.Range(CELLS["units"]).Value = 0
        excel.CalculateFull()
        poc0 = ws.Range(CELLS["poc"]).Value
        verdict0 = str(ws.Range(CELLS["verdict"]).Value)
        print(f"  units = 0 -> profit on cost {poc0}, verdict {verdict0!r}")
        failures += not isinstance(poc0, float)

        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()

    print("\n" + ("ALL CHECKS PASSED" if failures == 0 else f"{failures} CHECK(S) FAILED"))
    return 0 if failures == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
