# Design notes: rebuilding the budget app

This documents the redesign from the original `Budget Database 25.05_04.accdb`
(58 tables, mostly empty except reference data) into a smaller, automatable
schema, and *why* each change was made. Read this before importing anything.

## What was wrong with the old schema

1. **14 near-duplicate "fund snapshot" tables.** `TABLE_AACS`, `TABLE_AACS_Current`,
   `TABLE_F820Current`, `TABLE_F820Cumulative`, `TABLE_F826Current`,
   `TABLE_F826_CarryForward`, `TABLE_F826Undistributed`, `TABLE_F826_1321Balances`,
   `TABLE_SALT`, `TABLE_ALLW`, `TABLE_UpdatedSOA`, `TABLE_IFCAPbalances`,
   `TABLE_IFCAP_Details`, `TABLE_Rejects` are all "one row per FCP/fund with a
   dollar amount" shapes, each hardcoded to one VistA/FileMan report layout.
   Adding a new report type meant adding a new table. Replaced with one
   `FiscalSnapshots` fact table (see below) plus a table-driven import engine.

2. **5 near-identical pending-order tables**: `TABLE_Pending`,
   `TABLE_Pending_Updates`, `TABLE_Pending_Returned`, `TABLE_Unpaid_PC_Orders`,
   `TABLE_Unpaid_PC_Orders_PFY` all have the exact same columns. Collapsed into
   one `PendingOrders` table with a `RecordType` column.

3. **4 near-identical JV logbook tables**: `TABLE_JV_EB_LogBook`,
   `TABLE_JV_EW_LogBook`, `TABLE_JV_EBGIP_LogBook`, `TABLE_JV_EWGIP_LogBook`.
   The GIP variants were missing `ApprovedBy`/`Purpose`/`Type` columns the
   others had, for no apparent reason. Collapsed into one `JournalVouchers`
   table with a `LogBookCode` column, and every row now has an approval
   field (needed for the new alerts feature — "JV pending approval > N days").

4. **`TABLE_Vendors` is not a vendor table at all.** My first pass assumed it
   was a vendor registry missing a name column, and added `VendorName`. Then
   I actually looked at the 56 rows: `BOC` is unique across all of them, and
   there are only **7** distinct vendor numbers — 41 of the rows share the
   literal value `"MISCE"` (miscellaneous). It's a **BOC → default vendor
   number** lookup, keyed by BOC, not a list of vendors. Renamed to
   `BOCDefaultVendors(BOC PK, VendorNumber)`. There is no `VendorName` to
   fill in, because no vendor registry ever existed here.

   Consequence: `PendingOrders.VendorNumber` has **no** foreign key. There is
   nothing authoritative to point it at, and an FK against those 7 values
   would reject essentially every real purchase order.

5. **`TABLE_Attachments` stored file blobs inside the database**, which is
   most of why the file was 23 MB for basically no data. Replaced with a
   `FilePath` reference — point at a file share/SharePoint path instead of
   embedding the bytes.

6. **No real foreign keys anywhere** — `CostCenter`, `BOC`, `FCP` etc. were
   free-text/number columns repeated across tables with nothing enforcing
   they matched a real lookup row. The new schema adds actual FK constraints
   so bad imports fail loudly instead of silently producing orphan data.

   An FK is only added where the referenced table is genuinely the complete,
   authoritative list. `CostCenters` (2,019 rows) and `BudgetObjectCodes`
   (492) qualify. Two deliberately do **not** get FKs pointed at them:
   `BOCDefaultVendors` (not a vendor registry — see #4) and
   `StatusDescriptions` (its keys are truncated at 33 characters in the
   inherited data, e.g. `"Ordered (No Fiscal Action Require"`, so real
   status text would never match). Both remain useful for display lookups;
   they're just not constraints.

   `StatusDescriptions` also contained one duplicate key
   (`"Partial Payment (Partial Receipt)"`, two rows whose descriptions
   differed only in whitespace), which would have failed the new primary
   key outright. The seed CSV keeps one of them.

7. **Access/permission tables scattered across three tables**
   (`TABLE_FCP_DOAs`, `TABLE_IFCAP_Users`, `TABLE_FCP_SpecialAccess`) that are
   all "which person can touch which FCP." Collapsed into `Employees` +
   `FCPAccess`.

8. **No import history** — `TABLE_LastUpdated` was a single row that got
   overwritten on every refresh, so you could never see when something
   actually changed or whether an import silently failed. Replaced with
   `ImportLog` (one row per import run) and `ImportRejects` (bad rows tied to
   a specific run instead of dumped into an undifferentiated `TABLE_Rejects`).

## Tables intentionally dropped

`TABLE_Temp_Misc`, `TABLE_Temp_Misc1`, `TABLE_SavedTransactions`,
`TABLE_SavedPAIDQueries` — these had the shape of ad hoc query-result caches,
not real data. If it turns out one of these mattered, tell me what it was
for and I'll bring the concept back properly (as a saved query, not a table).

`TABLE_UserLog` — replaced by a lighter `AppLog` table used by the VBA error
handlers, since this is a single-user local app and the old per-login log
wasn't doing much.

`USysApplicationLog` — this is an Access-generated system table (data macro
error log). Access creates it itself; don't create it manually.

## Real file formats (F820, F20D, F826) — this superseded my first guess

The first version of this doc assumed F820/F826 downloads would be simple
delimited text (comma/caret-separated columns), and built a generic
column-position import engine (`modImport.bas`) around that assumption.
Once real sample files showed up in `Sample Files/`, that assumption turned
out to be wrong: these are **fixed-width printed reports** (VA FMS report
IDs `RBEACCVM`/`RBEACCV`/`RBESWSV`), not delimited data. They have repeating
page headers, multi-line section headers, summary/total rows mixed in with
detail rows, and — the part that actually breaks naive fixed-column-position
parsing — right-aligned dollar columns that overflow into the column to
their left when the number is wide enough (e.g. a `126,692,583.00` value
eating into where the account-name column would otherwise be).

Because of that overflow behavior, the parser doesn't slice at fixed column
positions for the dollar amounts. It finds them by pattern (rightmost
`-?[\d,]+\.\d\d`-shaped tokens on the line) and uses *that* position as the
right boundary for the text fields to its left, so a wide number never gets
misread as part of a cost-center code. This was worked out and validated
against the full sample files in Python first (0 parse errors across all
16,564 / 3,807 / 1,629 lines) before being ported to VBA in
`modImportReports.bas`, since I can't run/debug VBA directly from here.

Each report type's structure, briefly:

- **F820 / F20D** ("Monthly/Daily Activity by Account Classification Code"):
  one block per Account Classification Code, each block containing an FYTD
  Summary (Beginning/Ending Balance → `FiscalSnapshots`) and an itemized
  ledger of adjustments (Document ID, date, vendor, BOC, cost center, ceiling/
  obligation adjustment amounts → `FiscalTransactions`, a new table — this
  transaction-level detail didn't exist anywhere in the old schema at all).
  F820 is the monthly version, F20D is the daily version; same layout,
  different period.
- **F826** ("Status of Allowance"): one block per Fund/Program, each
  containing rows per Orgn/Act (essentially the same concept as F820's
  Account Classification Code, different report's name for it) with Budget/
  Obligations/Available amounts → `FiscalSnapshots`.

If a future download doesn't match — different RSD report ID, a column
that's missing, wider special-case dollar amounts — the bad lines land in
`ImportRejects` with the raw text and reason instead of silently mis-parsing.
Send me those and the parser gets adjusted.

One thing I couldn't verify from here: `CDate()` on `MM/DD/YY` values
assumes the machine running Access is set to US date format. That should be
true for a VA desktop, but if dates start importing as, e.g., day and month
swapped, that's the first thing to check.

## `FiscalSnapshots`: the one deliberately debatable choice

Instead of one rigid table per report (F820, F826, and whatever else shows
up later — AACS, SALT, ALLW...), there's one fact table:

```
FiscalSnapshots(SnapshotID, FileType, Station, BFYS, AO, FundCode, ClassCode,
                 ClassCodeName, Program, BalanceType, AmountType, Amount,
                 RunDate, AsOfDate, ImportID)
```

`FileType` says which report it came from (`F820`, `F20D`, `F826`, ...) and
`AmountType` says what the dollar figure means (`BudgetCeiling`,
`Obligations`, `UnobligatedBalance` for F820/F20D; `Budget`, `Obligations`,
`Available` for F826). This trades some type safety for the ability to
import a *new* report type by adding rows to `ImportFileTypes` and writing
one new parser sub in `modImportReports.bas` — no schema change needed. The
cost: reports that need a "wide" view (one row per class code with Ceiling/
Obligated/Available as separate columns) need a `PIVOT` query instead of a
plain `SELECT *`. See `qryF820EndingBalances` / `qryF826Status` in
`schema/sample_queries.sql`. If this turns out to be the wrong tradeoff in
practice — e.g. if a future report doesn't fit the "dimensions + one amount
per row" shape at all — a dedicated table for that report type is a fine
fallback, same as `FiscalTransactions` got one instead of being forced into
this fact table.

## What still has to be built by hand in Access

I can't run MS Access from this environment, so I can produce schema DDL and
VBA module source, but not the visual Form/Report designer output. See
`README.md` for exactly what to build and how the provided queries make the
forms mostly plumbing (bind a form to `qry*`, drop fields on it).
