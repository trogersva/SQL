Attribute VB_Name = "modJVLogbook"
Option Compare Database
Option Explicit

' Journal voucher / cost-transfer logbook automation, replacing the four
' separate EB/EW/EBGIP/EWGIP logbook tables with one JournalVouchers table
' (see DESIGN.md). Every entry now supports an approver, including the
' GIP variants that didn't have one before -- that's what makes
' qryUnapprovedJVs (modAlerts) possible.

' Generates the next DocID for a logbook, formatted <LogBookCode>-<FY>-<seq>,
' e.g. "EB-2026-0007". Sequence resets each fiscal year.
Public Function NextDocID(ByVal logBookCode As String, ByVal fy As String) As String
    Dim prefix As String
    prefix = logBookCode & "-" & fy & "-"

    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT DocID FROM JournalVouchers WHERE LogBookCode = '" & Replace(logBookCode, "'", "''") & _
        "' AND DocID Like '" & prefix & "*' ORDER BY DocID DESC")

    Dim nextSeq As Long
    nextSeq = 1
    If Not rs.EOF Then
        Dim lastDoc As String
        lastDoc = rs!DocID
        nextSeq = CLng(Mid(lastDoc, Len(prefix) + 1)) + 1
    End If
    rs.Close

    NextDocID = prefix & Format(nextSeq, "0000")
End Function

' Records a new JV. ApprovedByID is optional (Null = pending approval,
' which is what surfaces it in qryUnapprovedJVs).
Public Function InsertJV(ByVal logBookCode As String, ByVal transferTypeCode As String, _
    ByVal fy As String, ByVal shortDescription As String, ByVal purpose As String, _
    ByVal amount As Currency, ByVal enteredByID As Long, _
    Optional ByVal approvedByID As Variant = Null) As Long

    If amount = 0 Then
        Err.Raise vbObjectError + 10, , "JV amount cannot be zero."
    End If

    Dim docID As String
    docID = NextDocID(logBookCode, fy)

    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("JournalVouchers")
    rs.AddNew
    rs!docID = docID
    rs!TransferTypeCode = transferTypeCode
    rs!LogBookCode = logBookCode
    rs!TransDate = Now()
    rs!ShortDescription = shortDescription
    rs!Purpose = purpose
    rs!amount = amount
    rs!EnteredBy = enteredByID
    If Not IsNull(approvedByID) Then
        rs!ApprovedBy = approvedByID
        rs!ApprovedDate = Now()
    End If
    rs.Update
    rs.Bookmark = rs.LastModified
    InsertJV = rs!JVID
    rs.Close
End Function

' Approves an existing pending JV.
Public Sub ApproveJV(ByVal jvID As Long, ByVal approvedByID As Long)
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM JournalVouchers WHERE JVID = " & jvID)
    If rs.EOF Then
        Err.Raise vbObjectError + 11, , "JVID " & jvID & " not found."
    End If
    rs.Edit
    rs!ApprovedBy = approvedByID
    rs!ApprovedDate = Now()
    rs.Update
    rs.Close
End Sub
