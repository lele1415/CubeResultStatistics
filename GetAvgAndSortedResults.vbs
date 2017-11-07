'*************************************************
'****get avg and sorted results.
'*************************************************
Sub getSortedResult(sPureResult, iOptSeq, resultInfo)
    If iOptSeq = 17 Then
        resultInfo.BestResult = sPureResult
        Exit Sub
    End If

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

    Call compareToBestRecord(oOptInfo.BestRecord, resultInfo, iBestResult)
    If iOptSeq < 15 Then
        Call compareToAvgRecord(oOptInfo.AvgRecord, resultInfo, iAvgResult)
    End If

    resultInfo.BestResult = iBestResult
    resultInfo.AvgResult = iAvgResult
    resultInfo.SortedResults = joinArrayWithSpace(aPureResults)
End Sub

Sub compareToBestRecord(iBestRecord, resultInfo, iBestResult)
    If iBestResult - iBestRecord <= 0 Then
        resultInfo.IsBestBr = True
    Else
        resultInfo.IsBestBr = False
    End If
End Sub

Sub compareToAvgRecord(iAvgRecord, resultInfo, iAvgResult)
    If iAvgResult - iAvgRecord <= 0 Then
        resultInfo.IsAvgBr = True
    Else
        resultInfo.IsAvgBr = False
    End If
End Sub
