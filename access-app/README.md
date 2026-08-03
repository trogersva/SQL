# Budget App v2 (Access rebuild)

A redesigned replacement for `Budget Database 25.05_04.accdb`. Same tool
(VA fund-control-point budget tracking, purchase-card order tracking,
journal voucher logbook, VistA/FileMan data pulls), rebuilt as a normalized
schema plus VBA automation for imports, alerts, and reports, with everything
kept as text so it's actually version-controllable. See `docs/DESIGN.md` for
what changed and why.

This was built without access to MS Access itself (dev environment is
Linux), so the schema and VBA logic are ready to go, but the visual
Form/Report layer still needs to be built by hand in Access — see
"What's still manual" below.

## Setup (do this on your Windows machine with Access)

1. Create a new blank `.accdb` database.
2. Open the VBA editor (Alt+F11). For each file in `modules/`, right-click
   the project → **Import File...** and import it:
   - `modSchemaBuilder.bas`
   - `modImport.bas`
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
   into the matching table (append mode):
   - `CostCenters.csv` → `CostCenters`
   - `BudgetObjectCodes.csv` → `BudgetObjectCodes`
   - `StatusDescriptions.csv` → `StatusDescriptions`
   - `CostTransferTypes.csv` → `CostTransferTypes`
   - `Vendors.csv` → `Vendors` (this one only has `BOC`/`VendorNumber` from
     the old database — it has no vendor names, since the old table never
     captured them either. Fill in `VendorName` as you touch each vendor.)
   - `ImportFileTypes.csv` → `ImportFileTypes`
   - `ImportFieldMap.csv` → `ImportFieldMap`

   The last two are a **worked example only** (for the `Pending` orders
   file), not verified against a real export — I don't have a sample VistA/
   FileMan export file to confirm delimiter and column order. Send me one
   (redact dollar amounts/names if needed, I just need the structure) and
   I'll fill in accurate `ImportFileTypes`/`ImportFieldMap` rows for F820,
   F826, AACS, and the rest, instead of you guessing at the layout.

## Day to day use

**Importing a VistA/FileMan text export:**
```
ImportFile "Pending", "C:\path\to\export.txt"
```
Add a new report type by adding one row to `ImportFileTypes` and one row
per column to `ImportFieldMap` — no VBA changes needed. Rejected lines land
in `ImportRejects` tied to the `ImportLog` row, so a bad file never silently
drops data.

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
Dumps `qryFCPStatus`, `qryOverdueOrders`, `qryUnapprovedJVs`, `qryStaleImports`
to Excel files in `C:\BudgetAppReports\` (change `OutputFolder` at the top of
`modReports.bas` if you want them somewhere else, e.g. a synced SharePoint
folder).

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
