'*************************************************
'****get avg and sorted results.
'*************************************************
Sub getSortedResult(sPureResult, optInfo, resultInfo)
    Dim aTmp, aPureResults, iAvgResult, iBestResult, iEnd
    aPureResults = sortIntArray(Split(sPureResult))

    iBestResult = aPureResults(0)

    If optInfo.NeedNum = 5 Then
        iEnd = 3
    Else
        iEnd = 2
    End If

    If aPureResults(iEnd) <> 9999.99 Then
        iAvgResult = Round(((0 + aPureResults(iEnd-2) + aPureResults(iEnd-1) + aPureResults(iEnd)) / 3), 2)
        iAvgResult = FormatNumber(iAvgResult, 2, , , 0)
    Else
        iAvgResult = 9999.99
    End If

    Call compareToRecord(optInfo, resultInfo, iAvgResult, iBestResult)

    resultInfo.BestResult = iBestResult
    resultInfo.AvgResult = iAvgResult
    resultInfo.SortedResults = joinArrayWithSpace(aPureResults)
End Sub

Sub compareToRecord(optInfo, resultInfo, iAvgResult, iBestResult)
    Dim iAvgRecord, iBestRecord
    iBestRecord = optInfo.BestRecord
    iAvgRecord = optInfo.AvgRecord

    If optInfo.Seq <> OPT_SEQ_3fm Then
        iBestRecord = formatResultStr(iBestRecord)
        iAvgRecord = formatResultStr(iAvgRecord)
    End If

    If iBestResult - iBestRecord <= 0 Then resultInfo.IsBestBr = True
    If iAvgResult - iAvgRecord <= 0 Then resultInfo.IsAvgBr = True
End Sub
