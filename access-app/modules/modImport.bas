Attribute VB_Name = "modImport"
Option Compare Database
Option Explicit

' Generic, table-driven import engine for VistA/FileMan text exports.
'
' Setup once per file type (in ImportFileTypes / ImportFieldMap -- see
' seed/ImportFileTypes.csv and seed/ImportFieldMap.csv for a starting point):
'   - ImportFileTypes: one row per report (F820, F826, AACS, UnpaidPC, ...)
'     naming its delimiter, target table, and staleness threshold.
'   - ImportFieldMap: one row per column position in the source file.
'     TargetField is either a plain column name on the target table
'     (FCPNo, FY, Fund, DocID, VendorNumber, BOC, CostCenter, SnapshotDate...)
'     or, for FiscalSnapshots only, "Amount:<AmountType>" (e.g. "Amount:Ceiling")
'     which fans that one source column out into its own snapshot row --
'     this is how one F826 line with several dollar columns becomes several
'     FiscalSnapshots rows, one per AmountType. See DESIGN.md.
'
' Usage from the Immediate Window or a button:
'   ImportFile "F826", "C:\Exports\f826_2026q3.txt"

Public Sub ImportFile(ByVal fileType As String, ByVal filePath As String)
    Dim db As DAO.Database
    Set db = CurrentDb

    Dim cfg As DAO.Recordset
    Set cfg = db.OpenRecordset("SELECT * FROM ImportFileTypes WHERE FileType = '" & _
        Replace(fileType, "'", "''") & "'")
    If cfg.EOF Then
        MsgBox "No ImportFileTypes row for '" & fileType & "'. Add one first.", vbExclamation
        cfg.Close
        Exit Sub
    End If

    Dim delimiter As String, targetTable As String
    delimiter = Nz(cfg!delimiter, ",")
    targetTable = Nz(cfg!TargetTable, "")
    cfg.Close

    Dim fieldMap As DAO.Recordset
    Set fieldMap = db.OpenRecordset( _
        "SELECT * FROM ImportFieldMap WHERE FileType = '" & Replace(fileType, "'", "''") & _
        "' ORDER BY ColumnIndex")
    If fieldMap.EOF Then
        MsgBox "No ImportFieldMap rows for '" & fileType & "'. Add the column layout first.", vbExclamation
        fieldMap.Close
        Exit Sub
    End If

    Dim colTargets() As String, colTypes() As String
    Dim colCount As Long
    colCount = fieldMap.RecordCount
    ReDim colTargets(1 To colCount)
    ReDim colTypes(1 To colCount)
    Dim i As Long
    i = 1
    Do While Not fieldMap.EOF
        colTargets(i) = fieldMap!TargetField
        colTypes(i) = Nz(fieldMap!DataType, "TEXT")
        i = i + 1
        fieldMap.MoveNext
    Loop
    fieldMap.Close

    Dim importID As Long
    importID = StartImportLog(fileType, filePath)

    Dim rowsOK As Long, rowsBad As Long
    rowsOK = 0: rowsBad = 0

    Dim fnum As Integer, lineText As String, lineNo As Long
    fnum = FreeFile
    Open filePath For Input As #fnum
    lineNo = 0
    Do While Not EOF(fnum)
        Line Input #fnum, lineText
        lineNo = lineNo + 1
        If Len(Trim(lineText)) > 0 Then
            On Error GoTo LineFailed
            Dim parts() As String
            parts = Split(lineText, delimiter)
            If UBound(parts) - LBound(parts) + 1 <> colCount Then
                Err.Raise vbObjectError + 1, , "Expected " & colCount & " columns, got " & _
                    (UBound(parts) - LBound(parts) + 1)
            End If

            If targetTable = "FiscalSnapshots" Then
                InsertSnapshotRow db, fileType, importID, parts, colTargets, colTypes
            ElseIf targetTable = "PendingOrders" Then
                UpsertPendingOrderRow db, fileType, parts, colTargets, colTypes
            Else
                Err.Raise vbObjectError + 2, , "Unknown TargetTable: " & targetTable
            End If

            rowsOK = rowsOK + 1
            On Error GoTo 0
        End If
        GoTo NextLine

LineFailed:
        rowsBad = rowsBad + 1
        LogReject importID, lineNo, lineText, Err.Description
        Resume NextLine
NextLine:
    Loop
    Close #fnum

    FinishImportLog importID, rowsOK, rowsBad

    MsgBox "Import of '" & fileType & "' done: " & rowsOK & " rows loaded, " & _
           rowsBad & " rejected. See ImportRejects (ImportID=" & importID & ") for details.", _
           vbInformation
End Sub

Private Function StartImportLog(ByVal fileType As String, ByVal filePath As String) As Long
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("ImportLog")
    rs.AddNew
    rs!fileType = fileType
    rs!filePath = filePath
    rs!RowsImported = 0
    rs!RowsRejected = 0
    rs!ImportedAt = Now()
    rs!ImportedBy = Environ$("USERNAME")
    rs.Update
    rs.Bookmark = rs.LastModified
    StartImportLog = rs!ImportID
    rs.Close
End Function

Private Sub FinishImportLog(ByVal importID As Long, ByVal rowsOK As Long, ByVal rowsBad As Long)
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM ImportLog WHERE ImportID = " & importID)
    If Not rs.EOF Then
        rs.Edit
        rs!RowsImported = rowsOK
        rs!RowsRejected = rowsBad
        rs.Update
    End If
    rs.Close
End Sub

Private Sub LogReject(ByVal importID As Long, ByVal lineNo As Long, ByVal rawLine As String, ByVal msg As String)
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("ImportRejects")
    rs.AddNew
    rs!importID = importID
    rs!LineNumber = lineNo
    rs!RawLine = rawLine
    rs!ErrorMessage = Left(msg, 255)
    rs.Update
    rs.Close
End Sub

' Fans one source line out into one FiscalSnapshots row per "Amount:<Type>"
' column, sharing the plain dimension columns (FCPNo, FY, Fund, ...).
Private Sub InsertSnapshotRow(db As DAO.Database, ByVal fileType As String, ByVal importID As Long, _
    parts() As String, colTargets() As String, colTypes() As String)

    Dim dims As Object
    Set dims = CreateObject("Scripting.Dictionary")
    Dim amounts As Object
    Set amounts = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To UBound(colTargets)
        Dim target As String
        target = colTargets(i)
        If Left(target, 7) = "Amount:" Then
            amounts.Add Mid(target, 8), ConvertValue(parts(i - 1), colTypes(i))
        Else
            dims.Add target, ConvertValue(parts(i - 1), colTypes(i))
        End If
    Next i

    If amounts.Count = 0 Then
        Err.Raise vbObjectError + 3, , "No Amount:<Type> column mapped for " & fileType
    End If

    Dim amtType As Variant
    For Each amtType In amounts.Keys
        Dim rs As DAO.Recordset
        Set rs = db.OpenRecordset("FiscalSnapshots")
        rs.AddNew
        rs!fileType = fileType
        rs!importID = importID
        rs!AmountType = amtType
        rs!Amount = amounts(amtType)
        If dims.Exists("FCPNo") Then rs!FCPNo = dims("FCPNo")
        If dims.Exists("FY") Then rs!FY = dims("FY")
        If dims.Exists("Fund") Then rs!Fund = dims("Fund")
        If dims.Exists("DocID") Then rs!DocID = dims("DocID")
        If dims.Exists("VendorNumber") Then rs!VendorNumber = dims("VendorNumber")
        If dims.Exists("BOC") Then rs!BOC = dims("BOC")
        If dims.Exists("CostCenter") Then rs!CostCenter = dims("CostCenter")
        If dims.Exists("SnapshotDate") Then
            rs!SnapshotDate = dims("SnapshotDate")
        Else
            rs!SnapshotDate = Now()
        End If
        rs.Update
        rs.Close
    Next amtType
End Sub

Private Sub UpsertPendingOrderRow(db As DAO.Database, ByVal fileType As String, parts() As String, colTargets() As String, colTypes() As String)
    Dim vals As Object
    Set vals = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To UBound(colTargets)
        vals(colTargets(i)) = ConvertValue(parts(i - 1), colTypes(i))
    Next i

    ' RecordType isn't a column in the source file -- it's implied by which
    ' ImportFileTypes row you're importing (e.g. 'Pending' vs 'PFY' vs 'Returned').
    If Not vals.Exists("RecordType") Then
        vals.Add "RecordType", fileType
    End If

    If Not vals.Exists("TransactionNumber") Then
        Err.Raise vbObjectError + 4, , "PendingOrders import needs a TransactionNumber column"
    End If

    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM PendingOrders WHERE TransactionNumber = '" & _
        Replace(vals("TransactionNumber"), "'", "''") & "' AND RecordType = '" & _
        Replace(vals("RecordType"), "'", "''") & "'")

    If rs.EOF Then
        rs.AddNew
    Else
        rs.Edit
    End If

    Dim k As Variant
    For Each k In vals.Keys
        rs.Fields(CStr(k)).Value = vals(k)
    Next k
    rs!LastUpdated = Now()
    rs.Update
    rs.Close
End Sub

' NOTE: VistA/FileMan often exports dates in FileMan's own internal format
' (e.g. "3260803" for 2026-08-03, not a normal calendar string) rather than
' something CDate() understands. If DATE columns start rejecting rows, that's
' almost certainly why -- tell me the raw value from a real export and I'll
' add a FileMan-date converter here instead of guessing at the format now.
Private Function ConvertValue(ByVal raw As String, ByVal dataType As String) As Variant
    raw = Trim(raw)
    Select Case UCase(dataType)
        Case "CURRENCY"
            If Len(raw) = 0 Then
                ConvertValue = 0
            Else
                ConvertValue = CCur(raw)
            End If
        Case "LONG"
            If Len(raw) = 0 Then
                ConvertValue = Null
            Else
                ConvertValue = CLng(raw)
            End If
        Case "DATE"
            If Len(raw) = 0 Then
                ConvertValue = Null
            Else
                ConvertValue = CDate(raw)
            End If
        Case Else
            ConvertValue = raw
    End Select
End Function
