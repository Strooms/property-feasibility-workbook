# What survives the trip to Google Sheets

The workbook relies on three Excel features to do its job: locked cells, hidden
formulas, and conditional formatting. A brief that asks for a workbook which is
both locked down *and* works in Google Sheets is asking for two things that are
partly incompatible, so this is the test rather than an opinion.

**Method.** `output/property_feasibility.xlsx` uploaded to Google Drive and
opened in Google Sheets. Sheets kept it as `.xlsx` (Office editing mode — the
title bar shows an `.XLSX` badge) rather than converting to native Sheets
format. That is what happens when someone opens a distributed `.xlsx`, so it is
the realistic case.

---

## Result

| Feature | Excel | Google Sheets |
|---|---|---|
| Formulas calculate | ✓ | ✓ identical values — $1,246,038 and 28.3% both sides |
| Layout, fills, merges, fonts | ✓ | ✓ |
| Number formats | ✓ | ✓ |
| Verdict conditional formatting | ✓ | ✓ green banner rendered |
| Sensitivity colour scale | ✓ | ✓ full 5×5 grid rendered correctly |
| Data validation | ✓ | ✓ out-of-range entry rejected, cell unchanged |
| Hidden helper column | ✓ | ~ inconsistent (see below) |
| **In-cell data bars** | ✓ | **✗ gone** — values remain, bars do not render |
| **Formula hiding** | ✓ | **✗ formulas fully readable in the formula bar** |
| **Sheet protection** | ✓ | **✗ no protection at all** |

## The failure that matters

In Excel, typing into a formula cell is refused. In Google Sheets:

1. Clicked a locked, formula-hidden calculation cell. The formula bar showed
   `=K4*B15` in full.
2. Typed `999` and pressed Enter. **Accepted, with no warning.**
3. The model recalculated on the overwritten cell. Net profit went from
   $1,246,038 to $3,257,861. Profit on cost went from 28.3% to **136.5%**.
4. The verdict banner still read **MEETS THRESHOLD**, in green.

That is the exact failure the locking exists to prevent: not a visible error,
but a confident wrong answer. A lead magnet is handed to strangers, and the
first one who tab-keys through the sheet can silently break the model and take
its verdict at face value.

## The hidden helper column

The calculation chain lives in column K, hidden in Excel. On first open in
Sheets every intermediate value was visible down the right-hand side. After a
later reload it was hidden again. Inconsistent enough that it cannot be relied
on either way — which, given formulas are readable regardless, is mostly moot.

## Why formula hiding cannot be fixed

This is not an export bug to work around. Google Sheets has no equivalent of
Excel's formula-hidden attribute. Sheets protects *ranges* — it can warn or
block edits — but a viewer can always click a cell and read its formula. Any
version of this model that runs natively in Sheets shows its workings.

Protection itself *is* rebuildable in Sheets, as protected ranges, but that has
to be authored in Sheets. It does not arrive with the `.xlsx`.

## So the requirement has to give somewhere

Three honest options, and it is the client's call:

1. **Excel is the deliverable.** Locked and formula-hidden, as built. The Sheets
   version is best-effort: it calculates and looks right, but shows its formulas
   and cannot be protected.
2. **Both behave identically.** Drop formula hiding everywhere, rely on
   protection alone in Excel, and accept that Sheets users can edit anything.
3. **Two versions, maintained separately.** An Excel build with locking and
   hiding, and a native Sheets build with protected ranges — roughly double the
   revision surface.

Option 1 is usually right for a lead magnet: most recipients open it in Excel,
and the Sheets users still get a working calculator.

## Not tested

- **Native conversion** (*File → Save as Google Sheets*) was not tested
  separately. Formula hiding is absent in Sheets either way, so that finding
  holds; protection and data bars might behave differently after a true
  conversion.
- Sheets mobile apps.
- Whether re-exporting from Sheets back to `.xlsx` restores anything.
