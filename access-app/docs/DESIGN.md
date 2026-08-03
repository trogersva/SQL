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

4. **`TABLE_Vendors` had no vendor name** — just `BOC` and `VendorNumber`.
   Fixed: `Vendors.VendorName` added.

5. **`TABLE_Attachments` stored file blobs inside the database**, which is
   most of why the file was 23 MB for basically no data. Replaced with a
   `FilePath` reference — point at a file share/SharePoint path instead of
   embedding the bytes.

6. **No real foreign keys anywhere** — `CostCenter`, `BOC`, `FCP` etc. were
   free-text/number columns repeated across tables with nothing enforcing
   they matched a real lookup row. The new schema adds actual FK constraints
   so bad imports fail loudly instead of silently producing orphan data.

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

## `FiscalSnapshots`: the one deliberately debatable choice

Instead of one rigid table per VistA report (F820, F826, AACS, SALT, ALLW,
IFCAP...), there's one fact table:

```
FiscalSnapshots(SnapshotID, FileType, FCPNo, FY, Fund, DocID, VendorNumber,
                 BOC, CostCenter, AmountType, Amount, SnapshotDate, ImportID)
```

`FileType` says which report it came from (`F820`, `F826`, `AACS`, ...) and
`AmountType` says what the dollar figure means (`Ceiling`, `Obligated`,
`Budget`, `Unobligated`, `Uncommitted`, `Pending`, ...). This trades some
type safety for the ability to import a *new* VistA report type by adding a
row to `ImportFileTypes`/`ImportFieldMap` instead of writing a new table +
new VBA module every time. The cost: reports that need a "wide" view (one
row per FCP with Ceiling/Obligated/Available as separate columns) need a
`PIVOT` or crosstab query instead of a plain `SELECT *`. Example is in
`schema/03_sample_queries.sql`. If this turns out to be the wrong tradeoff
in practice, the old one-table-per-report style is a legitimate alternative
— easy to revisit once we see real import files.

## What still has to be built by hand in Access

I can't run MS Access from this environment, so I can produce schema DDL and
VBA module source, but not the visual Form/Report designer output. See
`README.md` for exactly what to build and how the provided queries make the
forms mostly plumbing (bind a form to `qry*`, drop fields on it).
