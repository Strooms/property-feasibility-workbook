"""Turn the filled-in workbook into a client-ready feasibility report.

Reads the live calculated values out of Excel, draws the charts, and assembles a
DOCX with a PDF alongside it.

    python build_report.py --sheet Residential --pdf
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import date
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import win32com.client as win32  # noqa: E402
from docx import Document  # noqa: E402
from docx.enum.table import WD_TABLE_ALIGNMENT  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Cm, Pt, RGBColor  # noqa: E402

ROOT = Path(__file__).parent
OUT = ROOT / "output"
WB = OUT / "property_feasibility.xlsx"

NAVY = RGBColor(0x1F, 0x3A, 0x5F)
INK = RGBColor(0x1A, 0x1A, 0x1A)
MUTED = RGBColor(0x7A, 0x7A, 0x7A)
GOOD = "3A7D5D"
BAD = "B3261E"
# python-docx wants bare hex, matplotlib wants a leading hash. Keep both.
GOOD_HEX = "#3A7D5D"
BAD_HEX = "#C9483C"

plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 9,
    "axes.edgecolor": "#9aa4b0", "axes.labelcolor": "#333333",
    "axes.titlesize": 10, "axes.titleweight": "bold",
    "xtick.color": "#555555", "ytick.color": "#555555",
})


# --- read ------------------------------------------------------------------


def read_values(sheet: str) -> dict:
    cells = json.loads((OUT / "cellmap.json").read_text(encoding="utf-8"))[sheet]
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(WB))
        ws = wb.Worksheets(sheet)
        excel.CalculateFull()

        def val(ref):
            return ws.Range(ref).Value

        data = {k: val(v) for k, v in cells.items()
                if isinstance(v, str) and k not in ("verdict",)}
        data["verdict"] = str(val(cells["verdict"]))
        data["land"] = [val(r) for r in cells["land"]]
        data["retail"] = [val(r) for r in cells["retail"]]

        shifts = cells["sensitivity"]["shifts"]
        top_left = cells["sensitivity"]["top_left"]
        col0 = "".join(ch for ch in top_left if ch.isalpha())
        row0 = int("".join(ch for ch in top_left if ch.isdigit()))
        grid = []
        for i in range(len(shifts)):
            row = []
            for j in range(len(shifts)):
                col = chr(ord(col0) + j)
                row.append(val(f"{col}{row0 + i}"))
            grid.append(row)
        data["shifts"] = shifts
        data["grid"] = grid

        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()
    return data


# --- charts ----------------------------------------------------------------


def chart_costs(d: dict, path: Path) -> Path:
    items = [
        ("Land + stamp duty", d["total_purchase"]),
        ("Build", d["build_total"]),
        ("Council contributions", d["council_total"]),
        ("Professional fees", d["prof_amt"]),
        ("Contingency", d["cont_amt"]),
        ("Finance", d["fin_amt"]),
    ][::-1]
    labels = [i[0] for i in items]
    values = [i[1] for i in items]
    total = d["total_cost"]

    fig, ax = plt.subplots(figsize=(6.2, 2.9))
    bars = ax.barh(labels, values, color="#31527a", height=0.62)
    ax.set_xlim(0, max(values) * 1.28)
    ax.spines[["top", "right"]].set_visible(False)
    ax.set_xlabel("Cost")
    ax.xaxis.set_major_formatter(lambda v, _: f"${v/1e6:.1f}M" if v else "0")
    for bar, v in zip(bars, values):
        ax.text(bar.get_width() + max(values) * 0.02,
                bar.get_y() + bar.get_height() / 2,
                f"${v:,.0f}   {v/total:.0%}", va="center", fontsize=8)
    ax.set_title("Where the money goes")
    fig.tight_layout()
    fig.savefig(path, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


def chart_sensitivity(d: dict, path: Path) -> Path:
    grid = d["grid"]
    shifts = d["shifts"]
    thr = d["threshold"]

    fig, ax = plt.subplots(figsize=(5.6, 3.2))
    cmap = matplotlib.colors.LinearSegmentedColormap.from_list(
        "poc", [BAD_HEX, "#F2C14E", GOOD_HEX])
    norm = matplotlib.colors.TwoSlopeNorm(vmin=min(min(r) for r in grid),
                                          vcenter=thr,
                                          vmax=max(max(r) for r in grid))
    ax.imshow(grid, cmap=cmap, norm=norm, aspect="auto")

    ax.set_xticks(range(len(shifts)), [f"{s:+.0%}" for s in shifts])
    ax.set_yticks(range(len(shifts)), [f"{s:+.0%}" for s in shifts])
    ax.set_xlabel("Build cost shift")
    ax.set_ylabel("Retail value shift")
    ax.set_title(f"Profit on cost  ·  threshold {thr:.0%}")
    for i, row in enumerate(grid):
        for j, v in enumerate(row):
            ax.text(j, i, f"{v:.1%}", ha="center", va="center", fontsize=8,
                    color="white" if abs(v - thr) > 0.12 else "#222222",
                    fontweight="bold" if v >= thr else "normal")
    ax.set_xticks([x - 0.5 for x in range(1, len(shifts))], minor=True)
    ax.set_yticks([y - 0.5 for y in range(1, len(shifts))], minor=True)
    ax.grid(which="minor", color="white", linewidth=1.5)
    ax.tick_params(which="minor", length=0)
    fig.tight_layout()
    fig.savefig(path, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


def chart_margin(d: dict, path: Path) -> Path:
    """A bullet bar: achieved profit on cost against the threshold."""
    poc, thr = d["poc"], d["threshold"]
    fig, ax = plt.subplots(figsize=(6.2, 1.15))
    top = max(poc, thr) * 1.45 if max(poc, thr) > 0 else 0.4
    low = min(0, poc * 1.2)
    ax.barh([0], [top - low], left=low, color="#eef1f5", height=0.55)
    ax.barh([0], [poc - low], left=low, height=0.55,
            color=GOOD_HEX if poc >= thr else BAD_HEX)
    ax.axvline(thr, color="#1a1a1a", linewidth=2)
    ax.text(thr, 0.42, f" threshold {thr:.0%}", fontsize=8, va="bottom")
    ax.text(poc, -0.45, f"{poc:.1%}", fontsize=11, fontweight="bold",
            ha="center", va="top",
            color=GOOD_HEX if poc >= thr else BAD_HEX)
    ax.set_xlim(low, top)
    ax.set_ylim(-0.9, 0.9)
    ax.set_yticks([])
    ax.xaxis.set_major_formatter(lambda v, _: f"{v:.0%}")
    ax.spines[["top", "right", "left"]].set_visible(False)
    fig.tight_layout()
    fig.savefig(path, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


# --- docx helpers ----------------------------------------------------------


def shade(cell, hex_fill):
    pr = cell._tc.get_or_add_tcPr()
    el = OxmlElement("w:shd")
    el.set(qn("w:val"), "clear")
    el.set(qn("w:fill"), hex_fill)
    pr.append(el)


def cell_text(cell, value, *, bold=False, size=9, colour=INK, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    if align is not None:
        p.alignment = align
    run = p.add_run(value)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.color.rgb = colour


def heading(doc, t, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16 if level == 1 else 10)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(t)
    r.bold = True
    r.font.size = Pt(13 if level == 1 else 10.5)
    r.font.color.rgb = NAVY if level == 1 else INK


def body(doc, t, *, size=9.5, colour=INK, italic=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(t)
    r.font.size = Pt(size)
    r.font.color.rgb = colour
    r.italic = italic


def kv_table(doc, rows, widths=(8.0, 5.0)):
    t = doc.add_table(rows=0, cols=2)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    for label, value in rows:
        cells = t.add_row().cells
        cells[0].width = Cm(widths[0])
        cells[1].width = Cm(widths[1])
        shade(cells[0], "F4F6F9")
        cell_text(cells[0], label, size=9)
        cell_text(cells[1], value, size=9, bold=True,
                  align=WD_ALIGN_PARAGRAPH.RIGHT)
    return t


def page_field(p):
    for instr, lit in (("PAGE", "1"), ("NUMPAGES", "1")):
        if instr == "NUMPAGES":
            p.add_run(" of ").font.size = Pt(8)
        begin = OxmlElement("w:fldChar"); begin.set(qn("w:fldCharType"), "begin")
        it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve")
        it.text = f" {instr} "
        sep = OxmlElement("w:fldChar"); sep.set(qn("w:fldCharType"), "separate")
        tx = OxmlElement("w:t"); tx.text = lit
        end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end")
        run = p.add_run(); run.font.size = Pt(8)
        for el in (begin, it, sep, tx, end):
            run._r.append(el)


# --- report ----------------------------------------------------------------


def build(d: dict, sheet: str, meta: dict, charts: dict, out: Path) -> Path:
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(9.5)

    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21.0), Cm(29.7)
    s.top_margin = s.bottom_margin = Cm(1.8)
    s.left_margin = s.right_margin = Cm(2.0)

    hdr = s.header.paragraphs[0]
    hdr.text = f"{sheet} Feasibility  ·  {meta['reference']}"
    hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for r in hdr.runs:
        r.font.size = Pt(8)
        r.font.color.rgb = MUTED
    ftr = s.footer.paragraphs[0]
    ftr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    page_field(ftr)

    meets = d["poc"] >= d["threshold"]

    # --- headline ---
    p = doc.add_paragraph()
    r = p.add_run(f"{sheet} Development")
    r.bold = True
    r.font.size = Pt(22)
    r.font.color.rgb = NAVY
    body(doc, f"Feasibility assessment  ·  {meta['issue_date']}", size=10,
         colour=MUTED, italic=True)

    v = doc.add_table(rows=1, cols=1)
    v.style = "Table Grid"
    cell = v.rows[0].cells[0]
    cell.width = Cm(17)
    shade(cell, "DDEEE2" if meets else "F6D5D2")
    cell_text(cell, d["verdict"], bold=True, size=18,
              colour=RGBColor.from_string(GOOD if meets else BAD),
              align=WD_ALIGN_PARAGRAPH.CENTER)

    body(doc, "")
    kv_table(doc, [
        ("Net profit", f"${d['net_profit']:,.0f}"),
        ("Profit on cost", f"{d['poc']:.1%}"),
        ("Threshold applied", f"{d['threshold']:.0%}"),
        ("Total cost", f"${d['total_cost']:,.0f}"),
        ("Gross realisation", f"${d['gross']:,.0f}"),
        ("Net realisation", f"${d['net_real']:,.0f}"),
    ])

    doc.add_picture(str(charts["margin"]), width=Cm(16.5))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Both figures are fractions, so the gap is too - convert before calling it
    # percentage points.
    margin_gap = (d["poc"] - d["threshold"]) * 100
    body(doc,
         f"The scheme returns {d['poc']:.1%} on cost against a {d['threshold']:.0%} "
         f"threshold — {abs(margin_gap):.1f} percentage points "
         f"{'above' if margin_gap >= 0 else 'below'} the bar. "
         f"Net profit is ${d['net_profit']:,.0f} on a total cost of "
         f"${d['total_cost']:,.0f}.")

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # --- the deal ---
    heading(doc, "1. The deal as entered")
    kv_table(doc, [
        ("Average raw land value (per lot)", f"${d['avg_land']:,.0f}"),
        ("Average retail value (per dwelling)", f"${d['avg_retail']:,.0f}"),
        ("Number of dwellings", f"{d['units']:,.0f}"),
        ("Build cost per dwelling", f"${d['build']:,.0f}"),
        ("Council contributions per dwelling", f"${d['council']:,.0f}"),
        ("Stamp duty", f"{d['stamp']:.1%}"),
        ("Professional fees", f"{d['prof']:.1%}"),
        ("Contingency", f"{d['cont']:.1%}"),
        ("Selling costs", f"{d['sell']:.1%}"),
        ("Finance costs", f"{d['fin']:.1%}"),
    ])

    heading(doc, "1.1 Comparables used", level=2)
    t = doc.add_table(rows=1, cols=3)
    t.style = "Table Grid"
    for i, h in enumerate(["#", "Raw land", "Retail"]):
        c = t.rows[0].cells[i]
        shade(c, "1F3A5F")
        cell_text(c, h, bold=True, size=9, colour=RGBColor(0xFF, 0xFF, 0xFF),
                  align=WD_ALIGN_PARAGRAPH.CENTER)
    for i, (land, retail) in enumerate(zip(d["land"], d["retail"]), start=1):
        cells = t.add_row().cells
        cell_text(cells[0], str(i), size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        cell_text(cells[1], f"${land:,.0f}", size=9, align=WD_ALIGN_PARAGRAPH.RIGHT)
        cell_text(cells[2], f"${retail:,.0f}", size=9, align=WD_ALIGN_PARAGRAPH.RIGHT)
    cells = t.add_row().cells
    shade(cells[0], "F4F6F9"); shade(cells[1], "F4F6F9"); shade(cells[2], "F4F6F9")
    cell_text(cells[0], "avg", bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(cells[1], f"${d['avg_land']:,.0f}", bold=True, size=9,
              align=WD_ALIGN_PARAGRAPH.RIGHT)
    cell_text(cells[2], f"${d['avg_retail']:,.0f}", bold=True, size=9,
              align=WD_ALIGN_PARAGRAPH.RIGHT)

    # --- costs ---
    heading(doc, "2. Where the money goes")
    doc.add_picture(str(charts["costs"]), width=Cm(16.5))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    body(doc,
         f"Land and stamp duty account for {d['total_purchase']/d['total_cost']:.0%} "
         f"of total cost, build for {d['build_total']/d['total_cost']:.0%}. "
         f"Together they are the only two lines large enough to change the verdict "
         f"on their own.", size=9, colour=MUTED, italic=True)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # --- sensitivity ---
    heading(doc, "3. What breaks it")
    body(doc,
         "Profit on cost if retail values and build costs move against the "
         "assumptions above. Bold figures clear the threshold.")
    doc.add_picture(str(charts["sensitivity"]), width=Cm(14.5))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    grid, shifts = d["grid"], d["shifts"]
    mid = len(shifts) // 2
    gdv_down = grid[mid - 1][mid]
    build_up = grid[mid][mid + 1]
    body(doc,
         f"A 5% fall in retail values takes the return to {gdv_down:.1%}. A 5% rise "
         f"in build cost takes it to {build_up:.1%}. The revenue side moves the "
         f"answer roughly twice as hard as the cost side, which is why comparable "
         f"selection deserves more scrutiny than the build estimate.")

    # --- method ---
    heading(doc, "4. Method and limits")
    body(doc,
         "Gross realisation is average retail value times dwelling count. Net "
         "realisation deducts selling costs. Total cost is land plus stamp duty "
         "plus build, council contributions, professional fees, contingency and "
         "finance. Profit on cost is net profit divided by total cost, and the "
         "verdict tests it against the threshold entered on the sheet.")
    body(doc, "This is a screening tool, not a full feasibility study. It does not "
              "model GST or the margin scheme, a project timeline, staged finance "
              "drawdown, demolition or site servicing, or strata establishment. "
              "Figures are only as good as the comparables entered.",
         size=9, colour=MUTED, italic=True)

    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    return out


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Build a feasibility report from the workbook.")
    ap.add_argument("--sheet", default="Residential", choices=["Residential", "Townhouse"])
    ap.add_argument("--client", default="Meridian Property Group")
    ap.add_argument("--pdf", action="store_true")
    args = ap.parse_args(argv)

    if not WB.exists():
        print(f"Workbook not found: {WB}  (run build_workbook.py first)")
        return 1

    print(f"> Reading   {WB.name}  [{args.sheet}]")
    d = read_values(args.sheet)
    print(f"  profit on cost {d['poc']:.2%}   verdict {d['verdict']}")

    print("> Charting")
    work = OUT / "_charts"
    work.mkdir(parents=True, exist_ok=True)
    charts = {
        "margin": chart_margin(d, work / f"margin_{args.sheet}.png"),
        "costs": chart_costs(d, work / f"costs_{args.sheet}.png"),
        "sensitivity": chart_sensitivity(d, work / f"sens_{args.sheet}.png"),
    }

    print("> Writing   DOCX")
    today = date.today()
    meta = {"client": args.client,
            "reference": f"FEAS-{today:%Y%m}-{args.sheet[:3].upper()}",
            "issue_date": f"{today:%d %B %Y}"}
    out = OUT / f"feasibility_report_{args.sheet.lower()}.docx"
    build(d, args.sheet, meta, charts, out)

    if args.pdf:
        print("> Writing   PDF")
        from docx2pdf import convert
        convert(str(out), str(out.with_suffix(".pdf")))

    print(f"\nDone\n  {out}")
    if args.pdf:
        print(f"  {out.with_suffix('.pdf')}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
