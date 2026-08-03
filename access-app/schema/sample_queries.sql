-- ============================================================
-- Reference queries against the new schema. modSchemaBuilder.bas
-- creates these as saved QueryDefs (qry*) so you can just open
-- them in Access, or bind a form/report to them directly.
-- ============================================================

-- qryF820EndingBalances: wide view of each fund/class-code's ending
-- balance for the reporting period, pivoting the FiscalSnapshots fact
-- table back into Budget Ceiling / Obligations / Unobligated columns.
-- Fed by modImportReports.ImportF820 / ImportF20D.
TRANSFORM Sum(Amount)
SELECT Station, BFYS, FundCode, ClassCode, ClassCodeName, RunDate
FROM FiscalSnapshots
WHERE FileType IN ('F820','F20D') AND BalanceType = 'Ending'
GROUP BY Station, BFYS, FundCode, ClassCode, ClassCodeName, RunDate
PIVOT AmountType;

-- qryF826Status: same idea for F826 (Status of Allowance), pivoted into
-- Budget / Obligations / Available columns per Orgn/Act.
-- Fed by modImportReports.ImportF826.
TRANSFORM Sum(Amount)
SELECT Station, BFYS, FundCode, Program, ClassCode, ClassCodeName, RunDate
FROM FiscalSnapshots
WHERE FileType = 'F826'
GROUP BY Station, BFYS, FundCode, Program, ClassCode, ClassCodeName, RunDate
PIVOT AmountType;

-- qryLargeAdjustments: individual F820/F20D ledger transactions above
-- $100k, in either direction. Adjust the threshold to taste.
SELECT FileType, ClassCode, ClassCodeName, DocID, TransDate, Vendor, BOC, CostCenter,
       CeilingAdjAmount, ObligationAdjAmount
FROM FiscalTransactions
WHERE Abs(CeilingAdjAmount) > 100000 OR Abs(ObligationAdjAmount) > 100000
ORDER BY TransDate DESC;

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
