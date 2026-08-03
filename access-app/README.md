# Budget App v2 (Access rebuild)

A redesigned replacement for `Budget Database 25.05_04.accdb`. Same tool
(VA fund-control-point budget tracking, purchase-card order tracking,
journal voucher logbook, F820/F20D/F826 VA FMS report imports), rebuilt as a
normalized schema plus VBA automation for imports, alerts, and reports, with
everything kept as text so it's actually version-controllable. See
`docs/DESIGN.md` for what changed and why.

This was built without access to MS Access itself (dev environment is
Linux), so the schema and VBA logic are ready to go, but the visual
Form/Report layer still needs to be built by hand in Access — see
"What's still manual" below.

## Setup (do this on your Windows machine with Access)

1. Create a new blank `.accdb` database.
2. Open the VBA editor (Alt+F11). For each file in `modules/`, right-click
   the project → **Import File...** and import it:
   - `modSchemaBuilder.bas`
   - `modImportCore.bas`
   - `modImport.bas`
   - `modImportReports.bas`
   - `modAlerts.bas`
   - `modJVLogbook.bas`
   - `modReports.bas`
3. In the VBA editor, **Tools → References**, make sure **Microsoft DAO
   3.6+ Object Library** is checked (Access usually has it on by default).
4. Open the Immediate Window (Ctrl+G) and run:
   ```
   SetupAll
   ```
   This creates all tables (`modSchemaBuilder.BuildSchema`) and the saved
   alert/report queries (`modSchemaBuilder.CreateSavedQueries`). Re-running
   it is safe — it skips tables that already exist.
5. Import the reference data in `seed/` via **External Data → Text File**
   into the matching table (append mode).

   **Wizard settings — the same for every file. Getting these wrong is the
   most likely thing to go visibly sideways:**

   | Wizard page | Setting |
   |---|---|
   | first page | **Append a copy of the records to the table:** *(pick the target table)* — not "Import into a new table" |
   | format | **Delimited** |
   | delimiter | **Comma** |
   | **Text Qualifier** | **`"`** ← must NOT be `{none}` |
   | | ☑ **First Row Contains Field Names** |

   Every field in every seed file is wrapped in double quotes, and those
   quotes are structural CSV syntax, not data. With Text Qualifier set to
   `"`, Access strips them. With it set to `{none}`, the quotes get stored
   as literal characters and you'll see `"120000"` sitting in the table —
   if you see that, this setting is why. Delete the imported rows, redo the
   import with the qualifier set, and they'll come in clean.

   The qualifier is genuinely required, not just cosmetic: several fields
   contain commas inside the value (cost-center and BOC descriptions, status
   descriptions), and in `ImportFileTypes.csv` the `Pending` row's
   `Delimiter` value *is* a literal comma. Without a qualifier those rows
   split into the wrong number of columns and the import fails or corrupts.

   Import in this order — `BudgetObjectCodes` must exist before
   `BOCDefaultVendors`, which has a foreign key to it:

   - `CostCenters.csv` → `CostCenters`
   - `BudgetObjectCodes.csv` → `BudgetObjectCodes`
   - `BOCDefaultVendors.csv` → `BOCDefaultVendors` (this is the old
     `TABLE_Vendors`, which turned out to be a BOC → default vendor-number
     lookup rather than a vendor list — see `docs/DESIGN.md`)
   - `StatusDescriptions.csv` → `StatusDescriptions`
   - `CostTransferTypes.csv` → `CostTransferTypes`
   - `ImportFileTypes.csv` → `ImportFileTypes` (do this one before importing
     F820/F20D/F826 data — `ImportLog` has a foreign key to it, so the first
     import will fail if these rows aren't there yet)
   - `ImportFieldMap.csv` → `ImportFieldMap` (only actually used by the
     `Pending` orders example below — unverified, see note there)

## Day to day use

**Importing F820 / F20D / F826 downloads:** these are fixed-width printed
reports, not delimited files, so they go through dedicated parsers rather
than the generic engine (see `docs/DESIGN.md` → "Real file formats"):
```
ImportF820 "C:\path\to\F820 0926.txt"
ImportF20D "C:\path\to\F20D - T_RBEACCV_073126.txt"
ImportF826 "C:\path\to\F826 07 31 26.txt"
```
The three files in `Sample Files/` at the repo root are good first-run test
data — the parser was validated against them line-by-line before being
written (zero parse errors across all of them). Try importing those first
to confirm your Access setup is wired up correctly before pointing it at a
fresh download. F820/F826 land in `FiscalSnapshots` (fund status balances)
and `FiscalTransactions` (the itemized adjustment ledger); F826 lands in
`FiscalSnapshots` only.

**Importing anything else that's genuinely a delimited file** (e.g. if
IFCAP ever hands you a comma/caret-separated pending-orders export instead
of another printed report):
```
ImportFile "Pending", "C:\path\to\export.txt"
```
This is the generic engine, driven by `ImportFileTypes`/`ImportFieldMap`.
The `Pending` row in the seed CSVs is an **unverified worked example** — I
don't have a real sample of this one. If it turns out IFCAP's pending-orders
export is actually another fixed-width printed report like F820, tell me
and I'll write it a dedicated parser the same way instead of forcing it
through here.

Either way, rejected lines land in `ImportRejects` tied to the `ImportLog`
row, so a bad file never silently drops data.

**Checking for anything that needs attention:**
```
ShowAlerts
```
Flags: data sources overdue for refresh, pending orders open >30 days,
journal vouchers awaiting approval, and due/overdue tasks. Wire this to run
automatically on open: **File → Options → Current Database → Display Form**,
or an `AutoExec` macro that calls `modAlerts.ShowAlerts`.

**Logging a journal voucher:**
```
InsertJV "EB", "EB-01", "2026", "Short desc", "Full purpose text", 1250.00, myEmployeeID
```
Auto-generates the DocID (`EB-2026-0001`, sequential per logbook per fiscal
year). Leave `approvedByID` off and it's flagged in `qryUnapprovedJVs` until
someone calls `ApproveJV`.

**Exporting the recurring reports:**
```
RunAllReports
```
Dumps `qryF820EndingBalances`, `qryF826Status`, `qryLargeAdjustments`,
`qryOverdueOrders`, `qryUnapprovedJVs`, `qryStaleImports` to Excel files in
`C:\BudgetAppReports\` (change `OutputFolder` at the top of `modReports.bas`
if you want them somewhere else, e.g. a synced SharePoint folder).

## What's still manual

I can't run the Access Form/Report designer from here, so these still need
to be built by hand in Access — but the queries do the hard part, so it's
mostly "new form → bind to this query → drop the fields on":

- A dashboard form bound to `qryStaleImports` / `qryOverdueOrders` /
  `qryUnapprovedJVs` / `qryDueTasks` (what `ShowAlerts` currently just pops
  a MsgBox summary of).
- Data entry forms for `PendingOrders`, `JournalVouchers`, `Tasks` — these
  can literally be Access's auto-generated form wizard bound to the table
  or a query, since the validation now lives in the VBA (`InsertJV`,
  `UpsertPendingOrderRow`) rather than needing form-level logic.
- Wiring `ShowAlerts` to run on startup (one AutoExec macro, see above).

Tell me once you've got a copy of Access open and I can walk through the
exact wizard steps for any of these, or write out the macro/VBA for a
specific form if you'd rather not use the wizard.
