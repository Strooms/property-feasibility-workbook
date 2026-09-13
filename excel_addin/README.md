# Generating the report from inside Excel

Two ways. Pick based on whether you want a button in Excel or not.

---

## Option A — no macros, no setup

Double-click **`Generate report.cmd`** in the project folder. It asks which tab,
runs the generator, and opens the PDF.

Nothing to install, nothing to trust, and it works on a machine where macros are
blocked by policy. This is the one to use unless you specifically want a button
on the ribbon.

---

## Option B — a button inside Excel

A small VBA add-in that adds **Generate feasibility report** to the ribbon. It
saves the workbook, runs the same Python generator, and opens the PDF.

### Install, once

1. Open Excel → a blank workbook → **Alt + F11** (Visual Basic editor).
2. **File → Import File…** → choose `FeasibilityReport.bas`.
3. Back in Excel: **File → Save As**, set *Save as type* to
   **Excel Add-in (\*.xlam)**, name it `FeasibilityReport.xlam`, and save it in
   the folder Excel offers by default (`%APPDATA%\Microsoft\AddIns`).
4. **File → Options → Add-ins → Manage: Excel Add-ins → Go…**, tick
   **FeasibilityReport**, click OK.
5. Right-click the ribbon → **Customise the Ribbon** → New Group on the Home
   tab → *Choose commands from:* **Macros** → add
   `GenerateFeasibilityReport` → rename it *Generate feasibility report*.

### Use it

1. Open `output/property_feasibility.xlsx`.
2. Go to the **Residential** or **Townhouse** tab.
3. Click the button. The PDF opens when it finishes, typically 30–40 seconds —
   most of that is Word starting up to do the PDF conversion.

### What it does

- Refuses politely if the workbook is unsaved, or if you are on the Start Here
  tab rather than a deal tab.
- Saves the workbook first, so the report reflects what is on screen rather than
  what was last written to disk.
- Looks for the project's own `.venv` and uses `pythonw.exe` from it, so no
  console window flashes up.
- Surfaces the generator's exit code if something fails, with the exact command
  to re-run in a terminal.

---

## Why this is *not* built into the workbook

The obvious version of this feature is a button embedded in the workbook itself,
which makes it a macro-enabled `.xlsm`. That is the wrong choice here, for three
reasons:

1. **Macros are blocked by default in downloaded files.** Windows marks files
   from the internet, and Office refuses to run their macros. A workbook
   distributed as a lead magnet would arrive with its button dead.
2. **Google Sheets cannot run VBA at all.** Any requirement that the workbook
   work in both Excel and Sheets is incompatible with embedded macros.
3. **It makes the file harder to trust.** A locked calculation workbook that
   also runs code is a bigger ask of the recipient than one that does not.

So the report generator lives beside the workbook rather than inside it. The
workbook stays a plain `.xlsx` that anyone can open anywhere; the button is a
tool for whoever produces the reports.

## Requirements

- Windows with Microsoft Excel (the generator reads live calculated values
  through Excel) and Word (for PDF export).
- The project's virtual environment, created as in the main README.
- The workbook must stay inside the project folder — the add-in locates
  `build_report.py` relative to the workbook's own path.
