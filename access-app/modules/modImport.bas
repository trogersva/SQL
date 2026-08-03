Attribute VB_Name = "modImport"
Option Compare Database
Option Explicit

' Generic, table-driven import engine for simple DELIMITED text exports
' (one row per data line, fixed column count). This does NOT handle F820,
' F20D, or F826 -- those are fixed-width printed reports with page headers
' and multi-line blocks, not delimited data. See modImportReports.bas for
' those. This engine is for a genuinely flat/delimited feed (e.g. if you
' get a comma- or caret-separated pending-orders export from IFCAP).
'
' Setup once per file type (in ImportFileTypes / ImportFieldMap -- see
' seed/ImportFileTypes.csv and seed/ImportFieldMap.csv for a starting point):
'   - ImportFileTypes: one row per feed, naming its delimiter and staleness
'     threshold. TargetTable must be 'PendingOrders'.
'   - ImportFieldMap: one row per column position in the source file,
'     TargetField naming a PendingOrders column (TransactionNumber, FCPNo,
'     PO, VendorNumber, ApprovalDate, Amount, StatusCode, ...).
'
' Usage from the Immediate Window or a button:
'   ImportFile "Pending", "C:\Exports\pending.txt"
'
' NOTE: this is an unverified starting point -- I don't have a real sample
' of a delimited pending-orders export to confirm layout against. If IFCAP
' actually hands you the same kind of fixed-width printed report as F820/
' F826/F20D instead, tell me and I'll write a dedicated parser for it in
' modImportReports.bas the same way, rather than forcing it through here.

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

    If targetTable <> "PendingOrders" Then
        MsgBox "modImport only supports TargetTable = 'PendingOrders'. " & _
               "F820/F20D/F826 go through modImportReports instead.", vbExclamation
        Exit Sub
    End If

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
    importID = modImportCore.StartImportLog(fileType, filePath)

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

            UpsertPendingOrderRow db, fileType, parts, colTargets, colTypes

            rowsOK = rowsOK + 1
            On Error GoTo 0
        End If
        GoTo NextLine

LineFailed:
        rowsBad = rowsBad + 1
        modImportCore.LogReject importID, lineNo, lineText, Err.Description
        Resume NextLine
NextLine:
    Loop
    Close #fnum

    modImportCore.FinishImportLog importID, rowsOK, rowsBad

    MsgBox "Import of '" & fileType & "' done: " & rowsOK & " rows loaded, " & _
           rowsBad & " rejected. See ImportRejects (ImportID=" & importID & ") for details.", _
           vbInformation
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
