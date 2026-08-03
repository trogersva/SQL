Attribute VB_Name = "modSchemaBuilder"
Option Compare Database
Option Explicit

' Creates the new schema in a blank .accdb. Run once from the
' Immediate Window (Ctrl+G): SetupAll
'
' Safe to re-run: every CREATE TABLE is skipped if the table
' already exists, and saved queries are dropped/recreated each run.

Public Sub SetupAll()
    BuildSchema
    CreateSavedQueries
    MsgBox "Schema build complete. Next: run modImport / import the" & vbCrLf & _
           "CSV files in /seed via External Data > Text File.", vbInformation
End Sub

Public Sub BuildSchema()
    Dim ddl() As String
    Dim i As Long

    ddl = TableDDL()

    For i = LBound(ddl) To UBound(ddl)
        CreateTableIfMissing ddl(i)
    Next i

    Debug.Print "BuildSchema: " & (UBound(ddl) - LBound(ddl) + 1) & " tables checked."
End Sub

Private Sub CreateTableIfMissing(ByVal ddl As String)
    Dim tableName As String
    tableName = ExtractTableName(ddl)

    If TableExists(tableName) Then
        Debug.Print "  skip (exists): " & tableName
        Exit Sub
    End If

    On Error GoTo Fail
    CurrentDb.Execute ddl, dbFailOnError
    Debug.Print "  created: " & tableName
    Exit Sub

Fail:
    Debug.Print "  FAILED: " & tableName & " -- " & Err.Number & ": " & Err.Description
    Debug.Print ddl
End Sub

Private Function ExtractTableName(ByVal ddl As String) As String
    Dim s As String, p1 As Long, p2 As Long
    s = Trim(ddl)
    p1 = InStr(1, s, "CREATE TABLE ", vbTextCompare) + Len("CREATE TABLE ")
    p2 = InStr(p1, s, "(")
    ExtractTableName = Trim(Left(Mid(s, p1), InStr(Mid(s, p1), "(") - 1))
End Function

Private Function TableExists(ByVal tableName As String) As Boolean
    Dim td As DAO.TableDef
    On Error Resume Next
    Set td = CurrentDb.TableDefs(tableName)
    TableExists = (Err.Number = 0) And Not (td Is Nothing)
    On Error GoTo 0
End Function

' The DDL below mirrors schema/schema.sql -- keep them in sync.
' s must be declared dynamic (Dim s() + ReDim) rather than Dim s(0 To 30):
' VBA refuses to ReDim a fixed-size array ("Array already dimensioned"), and
' the trailing ReDim Preserve trims it to however many entries were filled.
Private Function TableDDL() As String()
    Dim s() As String
    ReDim s(0 To 30)
    Dim n As Long
    n = 0

    s(n) = "CREATE TABLE CostCenters (" & _
        "CostCenter TEXT(10) NOT NULL, CCName TEXT(100), " & _
        "CONSTRAINT PK_CostCenters PRIMARY KEY (CostCenter))" : n = n + 1

    s(n) = "CREATE TABLE BudgetObjectCodes (" & _
        "BOC TEXT(10) NOT NULL, OC TEXT(4), BOCDescription TEXT(100), " & _
        "CONSTRAINT PK_BudgetObjectCodes PRIMARY KEY (BOC))" : n = n + 1

    ' Old TABLE_Vendors is a BOC -> default vendor-number lookup, not a
    ' vendor registry -- BOC is the unique key, VendorNumber repeats.
    s(n) = "CREATE TABLE BOCDefaultVendors (" & _
        "BOC TEXT(10) NOT NULL, VendorNumber TEXT(20), " & _
        "CONSTRAINT PK_BOCDefaultVendors PRIMARY KEY (BOC), " & _
        "CONSTRAINT FK_BOCDefaultVendors_BOC FOREIGN KEY (BOC) REFERENCES BudgetObjectCodes (BOC))" : n = n + 1

    s(n) = "CREATE TABLE StatusDescriptions (" & _
        "StatusCode TEXT(40) NOT NULL, Description MEMO, " & _
        "CONSTRAINT PK_StatusDescriptions PRIMARY KEY (StatusCode))" : n = n + 1

    s(n) = "CREATE TABLE CostTransferTypes (" & _
        "TransferTypeCode TEXT(20) NOT NULL, TransInitials TEXT(4), TransType TEXT(4), " & _
        "ExpRevGLBG TEXT(4), LogBookCode TEXT(10), " & _
        "CONSTRAINT PK_CostTransferTypes PRIMARY KEY (TransferTypeCode))" : n = n + 1

    s(n) = "CREATE TABLE Employees (" & _
        "EmployeeID AUTOINCREMENT, EmployeeName TEXT(100) NOT NULL, Title TEXT(100), " & _
        "Email TEXT(100), Initials TEXT(10), IsActive YESNO, " & _
        "CONSTRAINT PK_Employees PRIMARY KEY (EmployeeID))" : n = n + 1

    s(n) = "CREATE TABLE FundControlPoints (" & _
        "FCPNo LONG NOT NULL, FY TEXT(4) NOT NULL, ALLW TEXT(4), ACC TEXT(9), " & _
        "FCPName TEXT(100), BudgetID TEXT(22), Station TEXT(3), Fund TEXT(10), " & _
        "OrganizationalChart TEXT(100), ReportingCategory TEXT(100), " & _
        "Active YESNO, BudgetCall YESNO, FCPType TEXT(30), Notes MEMO, " & _
        "CONSTRAINT PK_FundControlPoints PRIMARY KEY (FCPNo, FY))" : n = n + 1

    s(n) = "CREATE TABLE FCPAccess (" & _
        "FCPAccessID AUTOINCREMENT, EmployeeID LONG NOT NULL, FCPNo LONG NOT NULL, " & _
        "FY TEXT(4) NOT NULL, AccessLevel TEXT(30), IsSpecialAccess YESNO, Notes TEXT(255), " & _
        "CONSTRAINT PK_FCPAccess PRIMARY KEY (FCPAccessID), " & _
        "CONSTRAINT FK_FCPAccess_Employee FOREIGN KEY (EmployeeID) REFERENCES Employees (EmployeeID), " & _
        "CONSTRAINT FK_FCPAccess_FCP FOREIGN KEY (FCPNo, FY) REFERENCES FundControlPoints (FCPNo, FY))" : n = n + 1

    s(n) = "CREATE TABLE ImportFileTypes (" & _
        "FileType TEXT(20) NOT NULL, DisplayName TEXT(100), Delimiter TEXT(5), " & _
        "TargetTable TEXT(60), StaleAfterDays INTEGER, " & _
        "CONSTRAINT PK_ImportFileTypes PRIMARY KEY (FileType))" : n = n + 1

    s(n) = "CREATE TABLE ImportFieldMap (" & _
        "FileType TEXT(20) NOT NULL, ColumnIndex INTEGER NOT NULL, TargetField TEXT(60) NOT NULL, " & _
        "DataType TEXT(20), " & _
        "CONSTRAINT PK_ImportFieldMap PRIMARY KEY (FileType, ColumnIndex), " & _
        "CONSTRAINT FK_ImportFieldMap_Type FOREIGN KEY (FileType) REFERENCES ImportFileTypes (FileType))" : n = n + 1

    s(n) = "CREATE TABLE ImportLog (" & _
        "ImportID AUTOINCREMENT, FileType TEXT(20), FilePath TEXT(255), " & _
        "RowsImported LONG, RowsRejected LONG, ImportedAt DATETIME, ImportedBy TEXT(100), " & _
        "CONSTRAINT PK_ImportLog PRIMARY KEY (ImportID), " & _
        "CONSTRAINT FK_ImportLog_Type FOREIGN KEY (FileType) REFERENCES ImportFileTypes (FileType))" : n = n + 1

    s(n) = "CREATE TABLE ImportRejects (" & _
        "RejectID AUTOINCREMENT, ImportID LONG, LineNumber LONG, RawLine MEMO, ErrorMessage TEXT(255), " & _
        "CONSTRAINT PK_ImportRejects PRIMARY KEY (RejectID), " & _
        "CONSTRAINT FK_ImportRejects_Import FOREIGN KEY (ImportID) REFERENCES ImportLog (ImportID))" : n = n + 1

    s(n) = "CREATE TABLE FiscalSnapshots (" & _
        "SnapshotID AUTOINCREMENT, FileType TEXT(20) NOT NULL, Station TEXT(3), BFYS TEXT(10), AO TEXT(4), " & _
        "FundCode TEXT(10), ClassCode TEXT(20), ClassCodeName TEXT(100), Program TEXT(10), BalanceType TEXT(10), " & _
        "AmountType TEXT(30) NOT NULL, Amount CURRENCY, RunDate DATETIME, AsOfDate DATETIME, ImportID LONG, " & _
        "CONSTRAINT PK_FiscalSnapshots PRIMARY KEY (SnapshotID), " & _
        "CONSTRAINT FK_FiscalSnapshots_Type FOREIGN KEY (FileType) REFERENCES ImportFileTypes (FileType), " & _
        "CONSTRAINT FK_FiscalSnapshots_Import FOREIGN KEY (ImportID) REFERENCES ImportLog (ImportID))" : n = n + 1

    s(n) = "CREATE TABLE FiscalTransactions (" & _
        "TransactionID AUTOINCREMENT, FileType TEXT(20) NOT NULL, Station TEXT(3), BFYS TEXT(10), AO TEXT(4), " & _
        "FundCode TEXT(10), ClassCode TEXT(20), ClassCodeName TEXT(100), DocID TEXT(20), TransDate DATETIME, " & _
        "Vendor TEXT(60), BOC TEXT(10), SubBOC TEXT(10), CostCenter TEXT(10), " & _
        "CeilingAdjAmount CURRENCY, ObligationAdjAmount CURRENCY, RunDate DATETIME, ImportID LONG, " & _
        "CONSTRAINT PK_FiscalTransactions PRIMARY KEY (TransactionID), " & _
        "CONSTRAINT FK_FiscalTransactions_Import FOREIGN KEY (ImportID) REFERENCES ImportLog (ImportID))" : n = n + 1

    s(n) = "CREATE TABLE PendingOrders (" & _
        "TransactionNumber TEXT(30) NOT NULL, RecordType TEXT(10) NOT NULL, FCPNo LONG, FY TEXT(4), " & _
        "PO TEXT(30), VendorNumber TEXT(20), BOC TEXT(10), CostCenter TEXT(10), " & _
        "ApprovalDate DATETIME, Amount CURRENCY, StatusCode TEXT(40), LastUpdated DATETIME, " & _
        "CONSTRAINT PK_PendingOrders PRIMARY KEY (TransactionNumber, RecordType), " & _
        "CONSTRAINT FK_PendingOrders_BOC FOREIGN KEY (BOC) REFERENCES BudgetObjectCodes (BOC), " & _
        "CONSTRAINT FK_PendingOrders_CC FOREIGN KEY (CostCenter) REFERENCES CostCenters (CostCenter))" : n = n + 1

    s(n) = "CREATE TABLE PendingAdjustments (" & _
        "AdjustmentID AUTOINCREMENT, TransactionNumber TEXT(30) NOT NULL, OriginalAmount CURRENCY, " & _
        "AdjustmentAmount CURRENCY, AdjustedTotal CURRENCY, ReversalDate DATETIME, Notes TEXT(255), " & _
        "CONSTRAINT PK_PendingAdjustments PRIMARY KEY (AdjustmentID))" : n = n + 1

    s(n) = "CREATE TABLE JournalVouchers (" & _
        "JVID AUTOINCREMENT, DocID TEXT(30), TransferTypeCode TEXT(20), LogBookCode TEXT(10) NOT NULL, " & _
        "TransDate DATETIME, ShortDescription TEXT(100), Purpose MEMO, Amount CURRENCY, " & _
        "EnteredBy LONG, ApprovedBy LONG, ApprovedDate DATETIME, " & _
        "CONSTRAINT PK_JournalVouchers PRIMARY KEY (JVID), " & _
        "CONSTRAINT FK_JV_TransferType FOREIGN KEY (TransferTypeCode) REFERENCES CostTransferTypes (TransferTypeCode), " & _
        "CONSTRAINT FK_JV_EnteredBy FOREIGN KEY (EnteredBy) REFERENCES Employees (EmployeeID), " & _
        "CONSTRAINT FK_JV_ApprovedBy FOREIGN KEY (ApprovedBy) REFERENCES Employees (EmployeeID))" : n = n + 1

    s(n) = "CREATE TABLE BudgetTransactions (" & _
        "TransID AUTOINCREMENT, TransDate DATETIME, FCPNo LONG, FY TEXT(4), TDA LONG, " & _
        "CostCenter TEXT(10), BOC TEXT(10), Description TEXT(150), " & _
        "Qtr1Amt CURRENCY, Qtr2Amt CURRENCY, Qtr3Amt CURRENCY, Qtr4Amt CURRENCY, " & _
        "IsRecurring YESNO, IsPaperOnly YESNO, IsFMSOnly YESNO, Notes MEMO, " & _
        "CONSTRAINT PK_BudgetTransactions PRIMARY KEY (TransID), " & _
        "CONSTRAINT FK_BudgetTrans_CC FOREIGN KEY (CostCenter) REFERENCES CostCenters (CostCenter), " & _
        "CONSTRAINT FK_BudgetTrans_BOC FOREIGN KEY (BOC) REFERENCES BudgetObjectCodes (BOC))" : n = n + 1

    s(n) = "CREATE TABLE BudgetAdjustments (" & _
        "AdjustmentID AUTOINCREMENT, AdjDate DATETIME, FCPNo LONG, FY TEXT(4), Qtr TEXT(4), " & _
        "CostCenter TEXT(10), BOC TEXT(10), AdjType TEXT(30), Amount CURRENCY, Description TEXT(150), " & _
        "CONSTRAINT PK_BudgetAdjustments PRIMARY KEY (AdjustmentID), " & _
        "CONSTRAINT FK_BudgetAdj_CC FOREIGN KEY (CostCenter) REFERENCES CostCenters (CostCenter), " & _
        "CONSTRAINT FK_BudgetAdj_BOC FOREIGN KEY (BOC) REFERENCES BudgetObjectCodes (BOC))" : n = n + 1

    s(n) = "CREATE TABLE Tasks (" & _
        "TaskID AUTOINCREMENT, TaskName TEXT(100) NOT NULL, Description MEMO, " & _
        "IsRecurring YESNO, RecurrenceDays INTEGER, DueDate DATETIME, CompletedDate DATETIME, CompletedBy LONG, " & _
        "CONSTRAINT PK_Tasks PRIMARY KEY (TaskID), " & _
        "CONSTRAINT FK_Tasks_CompletedBy FOREIGN KEY (CompletedBy) REFERENCES Employees (EmployeeID))" : n = n + 1

    s(n) = "CREATE TABLE Attachments (" & _
        "AttachmentID AUTOINCREMENT, Description TEXT(255), FilePath TEXT(255) NOT NULL, " & _
        "CONSTRAINT PK_Attachments PRIMARY KEY (AttachmentID))" : n = n + 1

    s(n) = "CREATE TABLE AppLog (" & _
        "LogID AUTOINCREMENT, LoggedAt DATETIME, Source TEXT(60), Message MEMO, " & _
        "CONSTRAINT PK_AppLog PRIMARY KEY (LogID))" : n = n + 1

    ReDim Preserve s(0 To n - 1)
    TableDDL = s
End Function

Public Sub CreateSavedQueries()
    DropQueryIfExists "qryF820EndingBalances"
    CurrentDb.CreateQueryDef "qryF820EndingBalances", _
        "TRANSFORM Sum(Amount) SELECT Station, BFYS, FundCode, ClassCode, ClassCodeName, RunDate " & _
        "FROM FiscalSnapshots WHERE FileType IN ('F820','F20D') AND BalanceType = 'Ending' " & _
        "GROUP BY Station, BFYS, FundCode, ClassCode, ClassCodeName, RunDate PIVOT AmountType;"

    DropQueryIfExists "qryF826Status"
    CurrentDb.CreateQueryDef "qryF826Status", _
        "TRANSFORM Sum(Amount) SELECT Station, BFYS, FundCode, Program, ClassCode, ClassCodeName, RunDate " & _
        "FROM FiscalSnapshots WHERE FileType = 'F826' " & _
        "GROUP BY Station, BFYS, FundCode, Program, ClassCode, ClassCodeName, RunDate PIVOT AmountType;"

    DropQueryIfExists "qryLargeAdjustments"
    CurrentDb.CreateQueryDef "qryLargeAdjustments", _
        "SELECT FileType, ClassCode, ClassCodeName, DocID, TransDate, Vendor, BOC, CostCenter, " & _
        "CeilingAdjAmount, ObligationAdjAmount " & _
        "FROM FiscalTransactions " & _
        "WHERE Abs(CeilingAdjAmount) > 100000 OR Abs(ObligationAdjAmount) > 100000 " & _
        "ORDER BY TransDate DESC;"

    DropQueryIfExists "qryStaleImports"
    CurrentDb.CreateQueryDef "qryStaleImports", _
        "SELECT t.FileType, t.DisplayName, t.StaleAfterDays, Max(l.ImportedAt) AS LastImported, " & _
        "DateDiff('d', Max(l.ImportedAt), Now()) AS DaysSinceImport " & _
        "FROM ImportFileTypes AS t LEFT JOIN ImportLog AS l ON l.FileType = t.FileType " & _
        "GROUP BY t.FileType, t.DisplayName, t.StaleAfterDays " & _
        "HAVING Max(l.ImportedAt) IS NULL OR DateDiff('d', Max(l.ImportedAt), Now()) > t.StaleAfterDays;"

    DropQueryIfExists "qryOverdueOrders"
    CurrentDb.CreateQueryDef "qryOverdueOrders", _
        "SELECT p.TransactionNumber, p.RecordType, p.FCPNo, p.Amount, p.ApprovalDate, " & _
        "DateDiff('d', p.ApprovalDate, Now()) AS DaysOpen, s.Description AS StatusDescription " & _
        "FROM PendingOrders AS p LEFT JOIN StatusDescriptions AS s ON s.StatusCode = p.StatusCode " & _
        "WHERE p.RecordType <> 'Returned' AND DateDiff('d', p.ApprovalDate, Now()) > 30;"

    DropQueryIfExists "qryUnapprovedJVs"
    CurrentDb.CreateQueryDef "qryUnapprovedJVs", _
        "SELECT jv.JVID, jv.DocID, jv.LogBookCode, jv.TransDate, jv.Amount, " & _
        "e.EmployeeName AS EnteredByName, DateDiff('d', jv.TransDate, Now()) AS DaysPending " & _
        "FROM JournalVouchers AS jv LEFT JOIN Employees AS e ON e.EmployeeID = jv.EnteredBy " & _
        "WHERE jv.ApprovedBy IS NULL;"

    DropQueryIfExists "qryDueTasks"
    CurrentDb.CreateQueryDef "qryDueTasks", _
        "SELECT TaskID, TaskName, DueDate, DateDiff('d', DueDate, Now()) AS DaysOverdue " & _
        "FROM Tasks WHERE CompletedDate IS NULL AND DueDate <= Now();"

    Debug.Print "CreateSavedQueries: done."
End Sub

Private Sub DropQueryIfExists(ByVal queryName As String)
    On Error Resume Next
    CurrentDb.QueryDefs.Delete queryName
    On Error GoTo 0
End Sub
