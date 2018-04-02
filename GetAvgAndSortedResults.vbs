'*************************************************
'****get avg and sorted results.
'*************************************************
Sub getSortedResult(sPureResult, iOptSeq, resultInfo)
    Dim oOptInfo
    Set oOptInfo = vaOptInfo.V(iOptSeq)

    If iOptSeq = OPT_SEQ_3mb Then
        Call getResultFor3mb(sPureResult, resultInfo)
        Call compareToBestRecord(iOptSeq, oOptInfo, resultInfo, resultInfo.BestResult, resultInfo.AvgResult)
        Exit Sub
    End If

    Dim aTmp, aPureResults, iAvgResult, iBestResult, iEnd
    aPureResults = sortIntArray(Split(sPureResult))
    iBestResult = aPureResults(0)

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

    Call compareToBestRecord(iOptSeq, oOptInfo, resultInfo, iBestResult, iAvgResult)
    Call compareToAvgRecord(iOptSeq, oOptInfo.AvgRecord, resultInfo, iAvgResult)

    resultInfo.BestResult = iBestResult
    resultInfo.AvgResult = iAvgResult
    resultInfo.SortedResults = joinArrayWithSpace(aPureResults)
End Sub

Sub compareToBestRecord(iOptSeq, oOptInfo, resultInfo, iBestResult, iAvgResult)
    Dim iBestRecord
    iBestRecord = oOptInfo.BestRecord

    If iOptSeq <> OPT_SEQ_3mb Then
        If iBestResult - iBestRecord <= 0 Then
            resultInfo.IsBestBr = True
        Else
            resultInfo.IsBestBr = False
        End If
    Else
        If iBestResult - iBestRecord > 0 Then
            resultInfo.IsBestBr = True
        ElseIf iBestResult - iBestRecord = 0 Then
            Dim iAvgRecord
            iAvgRecord = oOptInfo.AvgRecord
            If iAvgResult - iAvgRecord <= 0 Then
                resultInfo.IsBestBr = True
            Else
                resultInfo.IsBestBr = False
            End If
        Else
            resultInfo.IsBestBr = False
        End If
    End If
End Sub

Sub compareToAvgRecord(iOptSeq, iAvgRecord, resultInfo, iAvgResult)
    If iOptSeq >= 15 Then Exit Sub

    If iAvgResult - iAvgRecord <= 0 Then
        resultInfo.IsAvgBr = True
    Else
        resultInfo.IsAvgBr = False
    End If
End Sub

Sub getResultFor3mb(sPureResult, resultInfo)
    Dim aPureResults
    aPureResults = Split(sPureResult)
    resultInfo.BestResult = 2 * aPureResults(0) - aPureResults(1)
    resultInfo.AvgResult = aPureResults(2)
End Sub
