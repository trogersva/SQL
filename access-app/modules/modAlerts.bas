Attribute VB_Name = "modAlerts"
Option Compare Database
Option Explicit

' Reminders/alerts, built on the qryStaleImports / qryOverdueOrders /
' qryUnapprovedJVs / qryDueTasks saved queries from modSchemaBuilder.
'
' Wire ShowAlerts to run on startup: Access > Options > Current Database >
' Display Form, or an Autoexec macro that calls modAlerts.ShowAlerts.
' It only pops a summary -- open the underlying query for the row-level list.

Public Sub ShowAlerts()
    Dim msg As String
    Dim staleCount As Long, overdueCount As Long, unapprovedCount As Long, taskCount As Long

    staleCount = CountRows("qryStaleImports")
    overdueCount = CountRows("qryOverdueOrders")
    unapprovedCount = CountRows("qryUnapprovedJVs")
    taskCount = CountRows("qryDueTasks")

    If staleCount = 0 And overdueCount = 0 And unapprovedCount = 0 And taskCount = 0 Then
        Exit Sub ' nothing to say -- don't nag when there's nothing wrong
    End If

    msg = "Items needing attention:" & vbCrLf & vbCrLf
    If staleCount > 0 Then msg = msg & "- " & staleCount & " data source(s) overdue for refresh (qryStaleImports)" & vbCrLf
    If overdueCount > 0 Then msg = msg & "- " & overdueCount & " pending order(s) open > 30 days (qryOverdueOrders)" & vbCrLf
    If unapprovedCount > 0 Then msg = msg & "- " & unapprovedCount & " journal voucher(s) awaiting approval (qryUnapprovedJVs)" & vbCrLf
    If taskCount > 0 Then msg = msg & "- " & taskCount & " task(s) due or overdue (qryDueTasks)" & vbCrLf

    MsgBox msg, vbExclamation, "Budget App Alerts"
    LogAlert msg
End Sub

Private Function CountRows(ByVal queryName As String) As Long
    Dim rs As DAO.Recordset
    On Error GoTo NoQuery
    Set rs = CurrentDb.OpenRecordset(queryName)
    If rs.EOF Then
        CountRows = 0
    Else
        rs.MoveLast
        CountRows = rs.RecordCount
    End If
    rs.Close
    Exit Function
NoQuery:
    CountRows = 0
End Function

Private Sub LogAlert(ByVal msg As String)
    On Error Resume Next
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("AppLog")
    rs.AddNew
    rs!LoggedAt = Now()
    rs!Source = "modAlerts.ShowAlerts"
    rs!Message = msg
    rs.Update
    rs.Close
End Sub
