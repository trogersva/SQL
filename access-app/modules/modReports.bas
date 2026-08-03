Attribute VB_Name = "modReports"
Option Compare Database
Option Explicit

' Exports the recurring reports to Excel instead of running each query by
' hand. Point OutputFolder at wherever these need to land (a local folder,
' a synced OneDrive/SharePoint folder, etc).

Private Const OutputFolder As String = "C:\BudgetAppReports\"

Public Sub RunAllReports()
    EnsureFolder OutputFolder

    Dim stamp As String
    stamp = Format(Now(), "yyyy-mm-dd_hhnn")

    ExportQuery "qryFCPStatus", "FCPStatus_" & stamp & ".xlsx"
    ExportQuery "qryOverdueOrders", "OverdueOrders_" & stamp & ".xlsx"
    ExportQuery "qryUnapprovedJVs", "UnapprovedJVs_" & stamp & ".xlsx"
    ExportQuery "qryStaleImports", "StaleImports_" & stamp & ".xlsx"

    MsgBox "Reports exported to " & OutputFolder, vbInformation
End Sub

Private Sub ExportQuery(ByVal queryName As String, ByVal fileName As String)
    On Error GoTo Fail
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, _
        queryName, OutputFolder & fileName, True
    Exit Sub
Fail:
    LogReportFailure queryName, Err.Description
End Sub

Private Sub EnsureFolder(ByVal path As String)
    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If
End Sub

Private Sub LogReportFailure(ByVal queryName As String, ByVal msg As String)
    On Error Resume Next
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("AppLog")
    rs.AddNew
    rs!LoggedAt = Now()
    rs!Source = "modReports.RunAllReports"
    rs!Message = "Failed exporting " & queryName & ": " & msg
    rs.Update
    rs.Close
End Sub
