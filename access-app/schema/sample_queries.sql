-- ============================================================
-- Reference queries against the new schema. modSchemaBuilder.bas
-- creates these as saved QueryDefs (qry*) so you can just open
-- them in Access, or bind a form/report to them directly.
-- ============================================================

-- qryFCPStatus: wide view of an FCP's fund status, pivoting the
-- FiscalSnapshots fact table back into the old F826-style columns.
-- This is the tradeoff cost of the FiscalSnapshots design mentioned
-- in DESIGN.md -- one PIVOT instead of a plain SELECT *.
TRANSFORM Sum(Amount)
SELECT FCPNo, FY, Fund
FROM FiscalSnapshots
WHERE FileType = 'F826'
GROUP BY FCPNo, FY, Fund
PIVOT AmountType;

-- qryStaleImports: file types that haven't been refreshed within
-- their configured staleness window. Feeds the alerts module.
SELECT t.FileType, t.DisplayName, t.StaleAfterDays,
       Max(l.ImportedAt) AS LastImported,
       DateDiff('d', Max(l.ImportedAt), Now()) AS DaysSinceImport
FROM ImportFileTypes AS t
LEFT JOIN ImportLog AS l ON l.FileType = t.FileType
GROUP BY t.FileType, t.DisplayName, t.StaleAfterDays
HAVING Max(l.ImportedAt) IS NULL
    OR DateDiff('d', Max(l.ImportedAt), Now()) > t.StaleAfterDays;

-- qryOverdueOrders: pending orders sitting unresolved past 30 days
-- that aren't already in a terminal ("Paid"/"Returned") status.
SELECT p.TransactionNumber, p.RecordType, p.FCPNo, p.Amount,
       p.ApprovalDate, DateDiff('d', p.ApprovalDate, Now()) AS DaysOpen,
       s.Description AS StatusDescription
FROM PendingOrders AS p
LEFT JOIN StatusDescriptions AS s ON s.StatusCode = p.StatusCode
WHERE p.RecordType <> 'Returned'
  AND DateDiff('d', p.ApprovalDate, Now()) > 30;

-- qryUnapprovedJVs: journal vouchers still missing an approver.
SELECT jv.JVID, jv.DocID, jv.LogBookCode, jv.TransDate, jv.Amount,
       e.EmployeeName AS EnteredByName,
       DateDiff('d', jv.TransDate, Now()) AS DaysPending
FROM JournalVouchers AS jv
LEFT JOIN Employees AS e ON e.EmployeeID = jv.EnteredBy
WHERE jv.ApprovedBy IS NULL;

-- qryDueTasks: reminders due today or overdue and not completed.
SELECT TaskID, TaskName, DueDate, DateDiff('d', DueDate, Now()) AS DaysOverdue
FROM Tasks
WHERE CompletedDate IS NULL
  AND DueDate <= Now();
