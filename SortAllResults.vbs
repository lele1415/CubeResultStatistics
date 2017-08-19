'*************************************************
'****sort all results.
'*************************************************
Dim vaAllResults_333
Dim vaAllResults_444
Dim vaAllResults_555
Dim vaAllResults_222
Dim vaAllResults_3bf
Dim vaAllResults_3oh
Dim vaAllResults_3fm
Dim vaAllResults_mega
Dim vaAllResults_py
Dim vaAllResults_sq
Dim vaAllResults_clk
Dim vaAllResults_sk
Dim vaAllResults_666
Dim vaAllResults_777
Function sortAllResults()
    Call sortByOptName()

    Call sortByAvgResult(vaAllResults_333, true)
    Call sortByAvgResult(vaAllResults_444, true)
    Call sortByAvgResult(vaAllResults_555, true)
    Call sortByAvgResult(vaAllResults_222, true)
    Call sortByAvgResult(vaAllResults_3bf, false)
    Call sortByAvgResult(vaAllResults_3oh, true)
    Call sortByAvgResult(vaAllResults_3fm, true)
    Call sortByAvgResult(vaAllResults_mega, true)
    Call sortByAvgResult(vaAllResults_py, true)
    Call sortByAvgResult(vaAllResults_sq, true)
    Call sortByAvgResult(vaAllResults_clk, true)
    Call sortByAvgResult(vaAllResults_sk, true)
    Call sortByAvgResult(vaAllResults_666, true)
    Call sortByAvgResult(vaAllResults_777, true)

    MsgBox("OK")
End Function

    Function sortByOptName()
        Set vaAllResults_333 = New VariableArray
        Set vaAllResults_444 = New VariableArray
        Set vaAllResults_555 = New VariableArray
        Set vaAllResults_222 = New VariableArray
        Set vaAllResults_3bf = New VariableArray
        Set vaAllResults_3oh = New VariableArray
        Set vaAllResults_3fm = New VariableArray
        Set vaAllResults_mega = New VariableArray
        Set vaAllResults_py = New VariableArray
        Set vaAllResults_sq = New VariableArray
        Set vaAllResults_clk = New VariableArray
        Set vaAllResults_sk = New VariableArray
        Set vaAllResults_666 = New VariableArray
        Set vaAllResults_777 = New VariableArray

        Dim i, obj
        For i = 0 To vaAllValidResultInfo.Bound
            Set obj = vaAllValidResultInfo.V(i)
            Select Case obj.ResultOptSeq
                Case OPT_SEQ_333
                    vaAllResults_333.Append(obj)
                Case OPT_SEQ_444
                    vaAllResults_444.Append(obj)
                Case OPT_SEQ_555
                    vaAllResults_555.Append(obj)
                Case OPT_SEQ_222
                    vaAllResults_222.Append(obj)
                Case OPT_SEQ_3bf
                    vaAllResults_3bf.Append(obj)
                Case OPT_SEQ_3oh
                    vaAllResults_3oh.Append(obj)
                Case OPT_SEQ_3fm
                    vaAllResults_3fm.Append(obj)
                Case OPT_SEQ_mega
                    vaAllResults_mega.Append(obj)
                Case OPT_SEQ_py
                    vaAllResults_py.Append(obj)
                Case OPT_SEQ_sq
                    vaAllResults_sq.Append(obj)
                Case OPT_SEQ_clk
                    vaAllResults_clk.Append(obj)
                Case OPT_SEQ_sk
                    vaAllResults_sk.Append(obj)
                Case OPT_SEQ_666
                    vaAllResults_666.Append(obj)
                Case OPT_SEQ_777
                    vaAllResults_777.Append(obj)
            End Select
        Next
    End Function

    Function sortByAvgResult(vaObj, bCompartByAvg)
        If vaObj.Bound < 1 Then
            Exit Function
        End If

        Dim i, iCompareResult
        If bCompartByAvg Then
            For i = 0 To vaObj.Bound - 1
                iCompareResult = StrComp(vaObj.V(i).AvgResult, vaObj.V(i + 1).AvgResult)
                If iCompareResult > 0 Then
                    Call vaObj.SwapTwoValues(i, i + 1)
                ElseIf iCompareResult = 0 Then
                    iCompareResult = StrComp(vaObj.V(i).SortedResults, vaObj.V(i + 1).SortedResults)
                    If iCompareResult > 0 Then
                        Call vaObj.SwapTwoValues(i, i + 1)
                    End If
                End If
            Next
        Else
            For i = 0 To vaObj.Bound - 1
                iCompareResult = StrComp(vaObj.V(i).BestResult, vaObj.V(i + 1).BestResult)
                If iCompareResult > 0 Then
                    Call vaObj.SwapTwoValues(i, i + 1)
                ElseIf iCompareResult = 0 Then
                    iCompareResult = StrComp(vaObj.V(i).SortedResults, vaObj.V(i + 1).SortedResults)
                    If iCompareResult > 0 Then
                        Call vaObj.SwapTwoValues(i, i + 1)
                    End If
                End If
            Next
        End If
    End Function

