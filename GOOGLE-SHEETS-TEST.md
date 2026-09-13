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

## Can the formulas be hidden in Sheets at all?

Three approaches, all tested.

### 1. Apps Script custom function — hides the arithmetic, breaks Excel

Move the calculation into a script and call it from the cell. Tested on a native
Google Sheet:

```javascript
function NETREAL(land, retail, units) {
  var gross = retail * units;
  var net = gross - gross * 0.025;
  var purchase = land * units * 1.055;
  return net - purchase;
}
```

`=NETREAL(A1,B1,C1)` returned 3,746,250 — correct. And the formula bar shows
only `=NETREAL(A1,B1,C1)`. **The rates and the chain are genuinely not visible.**

Two costs, both material:

- **The file stops working in Excel.** Written into an `.xlsx` and opened in
  Excel, the same cell displays `#NAME?` — Excel has no such function. Verified,
  not assumed. For a brief whose primary platform is Excel, this is fatal.
- **It is obfuscation, not protection.** Anyone with edit access can open
  Extensions → Apps Script and read the code. The users of this workbook must
  type into it, so they have edit access by definition.

Worth having in the toolkit for Sheets-only work. Not an answer for this brief.

### 2. Keep the arithmetic in a hidden column — works, and keeps Excel

This is the architecture the workbook already uses, and it turns out to matter
more in Sheets than in Excel. The calculation chain lives in column K, hidden;
the dashboard cells are bare references.

Checked what each visible cell actually exposes:

| Cell | Formula a Sheets user can read |
|---|---|
| Net profit tile | `=K19` |
| Cost breakdown | `=K5` and similar |
| **Sensitivity grid** | `=IFERROR($K$11*(1+$D31)/(($K$8+($K$12*(1+E$30))+$K$13+…` |

So two thirds of the dashboard leaks nothing useful — a reference to a hidden
cell is not a formula anyone can learn from. **The sensitivity grid is the
exception**, and it currently carries the whole cost chain inline in all 25
cells.

Moving that arithmetic into the hidden block as well, leaving the grid cells as
bare references, would mean no visible cell in the workbook exposes a rate or a
relationship — in Excel *or* Sheets, with no scripts and no loss of
compatibility. That is a real design improvement rather than a workaround, and
it is roughly an hour's work.

It still is not security: a determined user can unhide the column. But it moves
the model from "readable at a glance" to "you have to go looking".

### 2b. How much does the protection actually buy? (Excel)

Worth knowing precisely, because it decides what can honestly be promised to a
client. Tested against the built workbook:

| Attempt | Result |
|---|---|
| Type into an input cell | accepted, no prompt |
| Type into a formula cell | refused |
| Read a hidden formula | refused — locked cells cannot even be selected |
| **Unhide column K** | **blocked** |
| `Review → Unprotect Sheet`, no password | one click, protection gone |
| Same, with `--password` set | rejected without the password |

So adding a password is worth doing, and it does **not** break the brief's
"no password prompt that blocks the user from typing in the inputs" — inputs are
unlocked, so nobody typing a deal is ever prompted. The password only stands
between a curious user and the Unprotect button. `build_workbook.py --password`
does this.

**But it is a deterrent, not a secret.** An `.xlsx` is a zip archive, and sheet
protection is a single XML element inside it. Deleting that element from the
three worksheets removed all protection in one pass, password or not, and the
formulas were readable again:

```
sheetProtection elements removed : 3
protection now active            : False
formula in K19 readable          : =K11-K18
```

That is not a flaw in the build; it is how the format works, and no Excel
workbook anywhere is better off. It sets what can be claimed:

> The workbook cannot be broken *by accident*, and its workings are not on
> display. It is not a vault, and nobody should be told it is one.

For a lead magnet that distinction barely matters — the risk being managed is a
prospect silently corrupting the model and trusting the answer, not a competitor
reverse-engineering it. Worth saying out loud to the client anyway.

### 3. Protected ranges — blocks edits, hides nothing

Sheets can restrict who edits a range, which fixes the overwrite problem from
the section above. It does not hide anything: the formula bar still shows the
formula to anyone who can select the cell. Has to be authored in Sheets; it does
not arrive with the `.xlsx`.

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

## Recommendation

Option 1 — Excel is the deliverable — **plus the hidden-column tightening from
§2 above**. That combination gives:

- Excel: locked, formula-hidden, protected. Unchanged.
- Sheets: calculates correctly, looks right, and exposes only bare references
  rather than the cost chain. Still editable, which has to be stated plainly to
  the client rather than glossed over.

No scripts, no second version to maintain, no loss of compatibility.

## Not tested

- **Native conversion** (*File → Save as Google Sheets*) was not tested
  separately. Formula hiding is absent in Sheets either way, so that finding
  holds; protection and data bars might behave differently after a true
  conversion.
- Whether a protected range in Sheets stops another editor unhiding a column.
  Verifying that needs a second account.
- Sheets mobile apps.
- Whether re-exporting from Sheets back to `.xlsx` restores anything.
