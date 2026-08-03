Attribute VB_Name = "modImportCore"
Option Compare Database
Option Explicit

' Shared logging helpers used by both modImport (generic delimited-file
' engine) and modImportReports (fixed-format F820/F20D/F826 parsers).

Public Function StartImportLog(ByVal fileType As String, ByVal filePath As String) As Long
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

Public Sub FinishImportLog(ByVal importID As Long, ByVal rowsOK As Long, ByVal rowsBad As Long)
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

Public Sub LogReject(ByVal importID As Long, ByVal lineNo As Long, ByVal rawLine As String, ByVal msg As String)
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

' Parses "1,234,567.89" / "-1,234.50" style numbers into a Currency value.
Public Function ParseAmount(ByVal raw As String) As Currency
    ParseAmount = CCur(Replace(Trim(raw), ",", ""))
End Function

' Parses MM/DD/YY as printed on these reports.
Public Function ParseReportDate(ByVal raw As String) As Variant
    raw = Trim(raw)
    If Len(raw) = 0 Then
        ParseReportDate = Null
    Else
        ParseReportDate = CDate(raw)
    End If
End Function
