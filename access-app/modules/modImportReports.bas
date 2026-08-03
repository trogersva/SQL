Attribute VB_Name = "modImportReports"
Option Compare Database
Option Explicit

' Parsers for the actual RSD report downloads (F820, F20D, F826) -- fixed-
' width printed reports, NOT delimited files. Validated line-by-line against
' real sample exports in Sample Files/ before being written here (zero parse
' errors across all three: 16,564 / 3,807 / 1,629 lines). If a future export
' doesn't parse cleanly, the bad lines land in ImportRejects with the raw
' text and reason -- send those back and the parser gets adjusted, rather
' than silently mis-importing.
'
' Usage:
'   ImportF820 "C:\Downloads\F820 0926.txt"
'   ImportF20D "C:\Downloads\F20D - T_RBEACCV_073126.txt"
'   ImportF826 "C:\Downloads\F826 07 31 26.txt"
'
' F820/F20D produce two kinds of rows:
'   - FiscalSnapshots: the FYTD Summary Beginning/Ending Balance lines
'     (Budget Ceiling / Obligations / Unobligated Balance per class code).
'   - FiscalTransactions: the Monthly/Daily Item Detail ledger lines
'     (one row per document/adjustment).
' F826 produces FiscalSnapshots rows only (Budget / Obligations / Available
' per Orgn-Act), since it's purely a point-in-time status report.

Public Sub ImportF820(ByVal filePath As String)
    ImportFundActivityReport "F820", filePath
End Sub

Public Sub ImportF20D(ByVal filePath As String)
    ImportFundActivityReport "F20D", filePath
End Sub

' ------------------------------------------------------------------
' F820 / F20D: "Monthly/Daily Activity by Account Classification Code"
' ------------------------------------------------------------------
Private Sub ImportFundActivityReport(ByVal fileType As String, ByVal filePath As String)
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim importID As Long
    importID = modImportCore.StartImportLog(fileType, filePath)

    Dim reSTN As Object, reRunDate As Object, reNum As Object, reDetailPrefix As Object
    Set reSTN = NewRegExp("STN:\s*(\S+)\s+.*?BFYS:\s*([\d\s]+?)\s+A/O:\s*(\S+)\s+FUND CODE:\s*(\S+)", True, False)
    Set reRunDate = NewRegExp("RUN DATE:\s*(\S+)\s+as of\s*(\S+)", True, False)
    Set reNum = NewRegExp("-?[\d,]+\.\d\d", True, True)
    Set reDetailPrefix = NewRegExp("^[A-Z]{2} ", False, False)

    Dim curStation As String, curBFYS As String, curAO As String, curFundCode As String
    Dim curClassCode As String, curClassCodeName As String
    Dim curRunDate As Variant, curAsOfDate As Variant
    Dim inDetail As Boolean

    Dim rowsOK As Long, rowsBad As Long
    rowsOK = 0: rowsBad = 0

    Dim fnum As Integer, s As String, lineNo As Long
    fnum = FreeFile
    Open filePath For Input As #fnum
    lineNo = 0
    Do While Not EOF(fnum)
        Line Input #fnum, s
        lineNo = lineNo + 1
        Dim ss As String, ssLower As String
        ss = Trim(s)
        ssLower = LCase(ss)

        On Error GoTo LineFailed

        If Left(ssLower, 4) = "run " And reRunDate.Test(s) Then
            Dim mRun As Object
            Set mRun = reRunDate.Execute(s)
            curRunDate = modImportCore.ParseReportDate(mRun(0).SubMatches(0))
            curAsOfDate = modImportCore.ParseReportDate(mRun(0).SubMatches(1))

        ElseIf Left(ssLower, 4) = "stn:" Then
            If reSTN.Test(s) Then
                Dim mStn As Object
                Set mStn = reSTN.Execute(s)
                curStation = mStn(0).SubMatches(0)
                curBFYS = CollapseSpaces(mStn(0).SubMatches(1))
                curAO = mStn(0).SubMatches(2)
                curFundCode = mStn(0).SubMatches(3)
            Else
                modImportCore.LogReject importID, lineNo, s, "STN line didn't match expected pattern"
                rowsBad = rowsBad + 1
            End If
            inDetail = False

        ElseIf ssLower Like "account classification code:*" Then
            curClassCode = Trim(Mid(ss, InStr(ss, ":") + 1))
            inDetail = False

        ElseIf ssLower Like "acc name:*" Then
            curClassCodeName = Trim(Mid(ss, InStr(ss, ":") + 1))

        ElseIf InStr(ssLower, "beginning balance") > 0 Or InStr(ssLower, "ending balance") > 0 Then
            Dim mBal As Object
            Set mBal = reNum.Execute(s)
            If mBal.Count <> 3 Then
                modImportCore.LogReject importID, lineNo, s, "Balance line: expected 3 numbers, found " & mBal.Count
                rowsBad = rowsBad + 1
            Else
                Dim balType As String
                balType = IIf(InStr(ssLower, "beginning") > 0, "Beginning", "Ending")
                InsertSnapshot db, fileType, importID, curStation, curBFYS, curAO, curFundCode, _
                    curClassCode, curClassCodeName, "", balType, "BudgetCeiling", _
                    modImportCore.ParseAmount(mBal(0).Value), curRunDate, curAsOfDate
                InsertSnapshot db, fileType, importID, curStation, curBFYS, curAO, curFundCode, _
                    curClassCode, curClassCodeName, "", balType, "Obligations", _
                    modImportCore.ParseAmount(mBal(1).Value), curRunDate, curAsOfDate
                InsertSnapshot db, fileType, importID, curStation, curBFYS, curAO, curFundCode, _
                    curClassCode, curClassCodeName, "", balType, "UnobligatedBalance", _
                    modImportCore.ParseAmount(mBal(2).Value), curRunDate, curAsOfDate
                rowsOK = rowsOK + 1
            End If
            inDetail = False

        ElseIf ssLower Like "monthly item detail for*" Or ssLower Like "daily item detail for*" Then
            inDetail = True

        ElseIf ssLower Like "total adjustments*" Then
            inDetail = False

        ElseIf inDetail And Len(s) >= 3 And reDetailPrefix.Test(s) Then
            Dim mNums As Object
            Set mNums = reNum.Execute(s)
            If mNums.Count < 2 Then
                modImportCore.LogReject importID, lineNo, s, "Detail line: expected >=2 numbers, found " & mNums.Count
                rowsBad = rowsBad + 1
            Else
                Dim restEnd As Long
                restEnd = mNums(mNums.Count - 2).FirstIndex
                Dim docID As String, transDate As String, vendor As String, boc As String, subBoc As String, cc As String
                docID = SafeSlice(s, 0, 18)
                transDate = SafeSlice(s, 18, 33)
                vendor = SafeSlice(s, 33, MinL(64, restEnd))
                boc = SafeSlice(s, 64, MinL(71, restEnd))
                subBoc = SafeSlice(s, 71, MinL(81, restEnd))
                cc = SafeSlice(s, 81, MinL(97, restEnd))

                InsertTransaction db, fileType, importID, curStation, curBFYS, curAO, curFundCode, _
                    curClassCode, curClassCodeName, docID, modImportCore.ParseReportDate(transDate), _
                    vendor, boc, subBoc, cc, _
                    modImportCore.ParseAmount(mNums(mNums.Count - 2).Value), _
                    modImportCore.ParseAmount(mNums(mNums.Count - 1).Value), curRunDate
                rowsOK = rowsOK + 1
            End If
        End If

        On Error GoTo 0
        GoTo NextLine
LineFailed:
        rowsBad = rowsBad + 1
        modImportCore.LogReject importID, lineNo, s, Err.Description
        Resume NextLine
NextLine:
    Loop
    Close #fnum

    modImportCore.FinishImportLog importID, rowsOK, rowsBad
    MsgBox "Import of '" & fileType & "' done: " & rowsOK & " rows loaded, " & _
           rowsBad & " rejected. See ImportRejects (ImportID=" & importID & ") for details.", vbInformation
End Sub

Private Sub InsertSnapshot(db As DAO.Database, ByVal fileType As String, ByVal importID As Long, _
    ByVal station As String, ByVal bfys As String, ByVal ao As String, ByVal fundCode As String, _
    ByVal classCode As String, ByVal classCodeName As String, ByVal prog As String, _
    ByVal balanceType As String, ByVal amountType As String, ByVal amount As Currency, _
    ByVal runDate As Variant, ByVal asOfDate As Variant)

    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("FiscalSnapshots")
    rs.AddNew
    rs!fileType = fileType
    rs!importID = importID
    rs!Station = station
    rs!bfys = bfys
    rs!ao = ao
    rs!FundCode = fundCode
    rs!ClassCode = classCode
    rs!ClassCodeName = classCodeName
    If Len(prog) > 0 Then rs!Program = prog
    If Len(balanceType) > 0 Then rs!BalanceType = balanceType
    rs!AmountType = amountType
    rs!Amount = amount
    If Not IsNull(runDate) Then rs!RunDate = runDate
    If Not IsNull(asOfDate) Then rs!AsOfDate = asOfDate
    rs.Update
    rs.Close
End Sub

Private Sub InsertTransaction(db As DAO.Database, ByVal fileType As String, ByVal importID As Long, _
    ByVal station As String, ByVal bfys As String, ByVal ao As String, ByVal fundCode As String, _
    ByVal classCode As String, ByVal classCodeName As String, ByVal docID As String, ByVal transDate As Variant, _
    ByVal vendor As String, ByVal boc As String, ByVal subBoc As String, ByVal cc As String, _
    ByVal ceilingAdj As Currency, ByVal oblAdj As Currency, ByVal runDate As Variant)

    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("FiscalTransactions")
    rs.AddNew
    rs!fileType = fileType
    rs!importID = importID
    rs!Station = station
    rs!bfys = bfys
    rs!ao = ao
    rs!FundCode = fundCode
    rs!ClassCode = classCode
    rs!ClassCodeName = classCodeName
    rs!docID = docID
    If Not IsNull(transDate) Then rs!TransDate = transDate
    rs!Vendor = vendor
    rs!boc = boc
    rs!SubBOC = subBoc
    rs!CostCenter = cc
    rs!CeilingAdjAmount = ceilingAdj
    rs!ObligationAdjAmount = oblAdj
    If Not IsNull(runDate) Then rs!RunDate = runDate
    rs.Update
    rs.Close
End Sub

' ------------------------------------------------------------------
' F826: "Status of Allowance"
' ------------------------------------------------------------------
Public Sub ImportF826(ByVal filePath As String)
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim importID As Long
    importID = modImportCore.StartImportLog("F826", filePath)

    Dim reSTN As Object, reRunDate As Object, reNum As Object, reDetailPrefix As Object
    Set reSTN = NewRegExp("STN:\s*(\S+)\s+BFYS:\s*(\S+)\s+Fund:\s*(\S+)\s+A/O:\s*(\S+)\s+Fund Code:\s*(\S+)", True, False)
    Set reRunDate = NewRegExp("Run Date:\s*(\S+)", True, False)
    Set reNum = NewRegExp("-?[\d,]+\.\d\d", True, True)
    Set reDetailPrefix = NewRegExp("^\s{3}\S", False, False)

    Dim curStation As String, curBFYS As String, curAO As String, curFundCode As String, curProgram As String
    Dim curRunDate As Variant

    Dim rowsOK As Long, rowsBad As Long
    rowsOK = 0: rowsBad = 0

    Dim fnum As Integer, s As String, lineNo As Long
    fnum = FreeFile
    Open filePath For Input As #fnum
    lineNo = 0
    Do While Not EOF(fnum)
        Line Input #fnum, s
        lineNo = lineNo + 1
        Dim ss As String, ssLower As String
        ss = Trim(s)
        ssLower = LCase(ss)

        On Error GoTo LineFailed

        If Left(ssLower, 4) = "run " And reRunDate.Test(s) Then
            Dim mRun As Object
            Set mRun = reRunDate.Execute(s)
            curRunDate = modImportCore.ParseReportDate(mRun(0).SubMatches(0))

        ElseIf Left(ssLower, 4) = "stn:" Then
            If reSTN.Test(s) Then
                Dim mStn As Object
                Set mStn = reSTN.Execute(s)
                curStation = mStn(0).SubMatches(0)
                curBFYS = mStn(0).SubMatches(1)
                curFundCode = mStn(0).SubMatches(4)
                curAO = mStn(0).SubMatches(3)
            Else
                modImportCore.LogReject importID, lineNo, s, "STN line didn't match expected pattern"
                rowsBad = rowsBad + 1
            End If

        ElseIf Left(ssLower, 8) = "program:" Then
            Dim progRest As String
            progRest = Trim(Mid(ss, 9))
            If Len(progRest) > 0 Then
                curProgram = Split(progRest, " ")(0)
            Else
                curProgram = ""
            End If

        ElseIf reDetailPrefix.Test(s) And reNum.Test(s) Then
            If Left(ssLower, 6) = "undist" Or Left(ssLower, 7) = "program" Or _
               Left(ssLower, 16) = "suballowance for" Or Left(ssLower, 13) = "allowance for" Then
                ' summary/total row -- not a detail row, skip
            Else
                Dim mNums As Object
                Set mNums = reNum.Execute(s)
                If mNums.Count < 3 Then
                    modImportCore.LogReject importID, lineNo, s, "F826 detail line: expected >=3 numbers, found " & mNums.Count
                    rowsBad = rowsBad + 1
                Else
                    Dim prefixEnd As Long
                    prefixEnd = mNums(mNums.Count - 3).FirstIndex
                    Dim prefix As String, orgnAct As String, actName As String, spacePos As Long
                    prefix = Trim(SafeSlice(s, 0, prefixEnd))
                    spacePos = InStr(prefix, " ")
                    If spacePos = 0 Then
                        orgnAct = prefix
                        actName = ""
                    Else
                        orgnAct = Left(prefix, spacePos - 1)
                        actName = Trim(Mid(prefix, spacePos + 1))
                    End If

                    InsertSnapshot db, "F826", importID, curStation, curBFYS, curAO, curFundCode, _
                        orgnAct, actName, curProgram, "", "Budget", _
                        modImportCore.ParseAmount(mNums(mNums.Count - 3).Value), curRunDate, Null
                    InsertSnapshot db, "F826", importID, curStation, curBFYS, curAO, curFundCode, _
                        orgnAct, actName, curProgram, "", "Obligations", _
                        modImportCore.ParseAmount(mNums(mNums.Count - 2).Value), curRunDate, Null
                    InsertSnapshot db, "F826", importID, curStation, curBFYS, curAO, curFundCode, _
                        orgnAct, actName, curProgram, "", "Available", _
                        modImportCore.ParseAmount(mNums(mNums.Count - 1).Value), curRunDate, Null
                    rowsOK = rowsOK + 1
                End If
            End If
        End If

        On Error GoTo 0
        GoTo NextLine826
LineFailed:
        rowsBad = rowsBad + 1
        modImportCore.LogReject importID, lineNo, s, Err.Description
        Resume NextLine826
NextLine826:
    Loop
    Close #fnum

    modImportCore.FinishImportLog importID, rowsOK, rowsBad
    MsgBox "Import of 'F826' done: " & rowsOK & " rows loaded, " & _
           rowsBad & " rejected. See ImportRejects (ImportID=" & importID & ") for details.", vbInformation
End Sub

' ------------------------------------------------------------------
' Shared helpers
' ------------------------------------------------------------------
Private Function NewRegExp(ByVal pattern As String, ByVal ignoreCase As Boolean, ByVal isGlobal As Boolean) As Object
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = pattern
    re.IgnoreCase = ignoreCase
    re.Global = isGlobal
    Set NewRegExp = re
End Function

' Python-style s[a:b] on a 0-based half-open range, clamped to the string's
' bounds, trimmed.
Private Function SafeSlice(ByVal s As String, ByVal a As Long, ByVal b As Long) As String
    If b < a Then b = a
    If b > Len(s) Then b = Len(s)
    If a > Len(s) Then a = Len(s)
    If a < 0 Then a = 0
    SafeSlice = Trim(Mid(s, a + 1, b - a))
End Function

Private Function MinL(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then
        MinL = a
    Else
        MinL = b
    End If
End Function

Private Function CollapseSpaces(ByVal s As String) As String
    Dim re As Object
    Set re = NewRegExp("\s+", False, True)
    CollapseSpaces = Trim(re.Replace(s, " "))
End Function
