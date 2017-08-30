'*************************************************
'****get avg and sorted results.
'*************************************************
Sub getSortedResult(sPureResult, iOptSeq, resultInfo)
    Dim aTmp, aPureResults, iAvgResult, iBestResult, oOptInfo, iEnd
    aPureResults = sortIntArray(Split(sPureResult))
    iBestResult = aPureResults(0)
    Set oOptInfo = vaOptInfo.V(iOptSeq)

    If oOptInfo.NeedNum = 5 Then
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

    Call compareToRecord(oOptInfo, resultInfo, iAvgResult, iBestResult)

    resultInfo.BestResult = iBestResult
    resultInfo.AvgResult = iAvgResult
    resultInfo.SortedResults = joinArrayWithSpace(aPureResults)
End Sub

Sub compareToRecord(oOptInfo, resultInfo, iAvgResult, iBestResult)
    Dim iAvgRecord, iBestRecord
    iBestRecord = oOptInfo.BestRecord
    iAvgRecord = oOptInfo.AvgRecord

    If iBestResult - iBestRecord <= 0 Then
        resultInfo.IsBestBr = True
    Else
        resultInfo.IsBestBr = False
    End If

    If iAvgResult - iAvgRecord <= 0 Then
        resultInfo.IsAvgBr = True
    Else
        resultInfo.IsAvgBr = False
    End If
End Sub
