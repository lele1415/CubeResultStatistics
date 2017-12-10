'*************************************************
'****sort all results.
'*************************************************
Const ID_SORT_RESULTS_STATUS = "sort_results_status"
Const ID_BUTTON_SAVE_ALL_RESULTS = "save_all_results"

Const SORT_BY_AVG = 0
Const SORT_BY_BEST = 1
Const SORT_BY_BEST_ANTI = 2

Dim vaAllResults_333
Dim vaAllResults_222
Dim vaAllResults_444
Dim vaAllResults_555
Dim vaAllResults_666
Dim vaAllResults_777
Dim vaAllResults_3bf
Dim vaAllResults_3fm
Dim vaAllResults_3oh
Dim vaAllResults_3wf
Dim vaAllResults_mega
Dim vaAllResults_py
Dim vaAllResults_clk
Dim vaAllResults_sk
Dim vaAllResults_sq
Dim vaAllResults_4bf
Dim vaAllResults_5bf
Dim vaAllResults_3mb

Sub onClickSortAllResults()
    Call setInnerHtml(ID_SORT_RESULTS_STATUS, "正在排序...")
    idTimer = window.setTimeout("sortAllResults()", 0, "VBScript")
End Sub

Sub sortAllResults()
    window.clearTimeout(idTimer)

    Call sortByOptName()

    Call sortByAvgResult(vaAllResults_333, SORT_BY_AVG)
    Call sortByAvgResult(vaAllResults_222, SORT_BY_AVG)
    Call sortByAvgResult(vaAllResults_444, SORT_BY_AVG)
    Call sortByAvgResult(vaAllResults_555, SORT_BY_AVG)
    Call sortByAvgResult(vaAllResults_666, SORT_BY_AVG)
    Call sortByAvgResult(vaAllResults_777, SORT_BY_AVG)
    Call sortByAvgResult(vaAllResults_3bf, SORT_BY_BEST)
    Call sortByAvgResult(vaAllResults_3fm, SORT_BY_AVG)
    Call sortByAvgResult(vaAllResults_3oh, SORT_BY_AVG)
    Call sortByAvgResult(vaAllResults_3wf, SORT_BY_AVG)
    Call sortByAvgResult(vaAllResults_mega, SORT_BY_AVG)
    Call sortByAvgResult(vaAllResults_py, SORT_BY_AVG)
    Call sortByAvgResult(vaAllResults_clk, SORT_BY_AVG)
    Call sortByAvgResult(vaAllResults_sk, SORT_BY_AVG)
    Call sortByAvgResult(vaAllResults_sq, SORT_BY_AVG)
    Call sortByAvgResult(vaAllResults_4bf, SORT_BY_BEST)
    Call sortByAvgResult(vaAllResults_5bf, SORT_BY_BEST)
    Call sortByAvgResult(vaAllResults_3mb, SORT_BY_BEST_ANTI)

    Call enableElementAfterSortAllResults()
    Call setInnerHtml(ID_SORT_RESULTS_STATUS, "已完成")
End Sub

    Sub sortByOptName()
        Set vaAllResults_333 = New VariableArray
        Set vaAllResults_222 = New VariableArray
        Set vaAllResults_444 = New VariableArray
        Set vaAllResults_555 = New VariableArray
        Set vaAllResults_666 = New VariableArray
        Set vaAllResults_777 = New VariableArray
        Set vaAllResults_3bf = New VariableArray
        Set vaAllResults_3fm = New VariableArray
        Set vaAllResults_3oh = New VariableArray
        Set vaAllResults_3wf = New VariableArray
        Set vaAllResults_mega = New VariableArray
        Set vaAllResults_py = New VariableArray
        Set vaAllResults_clk = New VariableArray
        Set vaAllResults_sk = New VariableArray
        Set vaAllResults_sq = New VariableArray
        Set vaAllResults_4bf = New VariableArray
        Set vaAllResults_5bf = New VariableArray
        Set vaAllResults_3mb = New VariableArray
        Call vaAllResults_333.SetPreBound(vaOptInfo.V(0).Count)
        Call vaAllResults_222.SetPreBound(vaOptInfo.V(1).Count)
        Call vaAllResults_444.SetPreBound(vaOptInfo.V(2).Count)
        Call vaAllResults_555.SetPreBound(vaOptInfo.V(3).Count)
        Call vaAllResults_666.SetPreBound(vaOptInfo.V(4).Count)
        Call vaAllResults_777.SetPreBound(vaOptInfo.V(5).Count)
        Call vaAllResults_3bf.SetPreBound(vaOptInfo.V(6).Count)
        Call vaAllResults_3fm.SetPreBound(vaOptInfo.V(7).Count)
        Call vaAllResults_3oh.SetPreBound(vaOptInfo.V(8).Count)
        Call vaAllResults_3wf.SetPreBound(vaOptInfo.V(9).Count)
        Call vaAllResults_mega.SetPreBound(vaOptInfo.V(10).Count)
        Call vaAllResults_py.SetPreBound(vaOptInfo.V(11).Count)
        Call vaAllResults_clk.SetPreBound(vaOptInfo.V(12).Count)
        Call vaAllResults_sk.SetPreBound(vaOptInfo.V(13).Count)
        Call vaAllResults_sq.SetPreBound(vaOptInfo.V(14).Count)
        Call vaAllResults_4bf.SetPreBound(vaOptInfo.V(15).Count)
        Call vaAllResults_5bf.SetPreBound(vaOptInfo.V(16).Count)
        Call vaAllResults_3mb.SetPreBound(vaOptInfo.V(17).Count)

        Dim i, resultObj
        For i = 0 To vaAllValidResultInfo.Bound
            Set resultObj = vaAllValidResultInfo.V(i)
            Select Case resultObj.ResultOptSeq
                Case OPT_SEQ_333
                    Call appendResultObj(vaAllResults_333, resultObj)
                Case OPT_SEQ_222
                    Call appendResultObj(vaAllResults_222, resultObj)
                Case OPT_SEQ_444
                    Call appendResultObj(vaAllResults_444, resultObj)
                Case OPT_SEQ_555
                    Call appendResultObj(vaAllResults_555, resultObj)
                Case OPT_SEQ_666
                    Call appendResultObj(vaAllResults_666, resultObj)
                Case OPT_SEQ_777
                    Call appendResultObj(vaAllResults_777, resultObj)
                Case OPT_SEQ_3bf
                    Call appendResultObj(vaAllResults_3bf, resultObj)
                Case OPT_SEQ_3fm
                    Call appendResultObj(vaAllResults_3fm, resultObj)
                Case OPT_SEQ_3oh
                    Call appendResultObj(vaAllResults_3oh, resultObj)
                Case OPT_SEQ_3wf
                    Call appendResultObj(vaAllResults_3wf, resultObj)
                Case OPT_SEQ_mega
                    Call appendResultObj(vaAllResults_mega, resultObj)
                Case OPT_SEQ_py
                    Call appendResultObj(vaAllResults_py, resultObj)
                Case OPT_SEQ_clk
                    Call appendResultObj(vaAllResults_clk, resultObj)
                Case OPT_SEQ_sk
                    Call appendResultObj(vaAllResults_sk, resultObj)
                Case OPT_SEQ_sq
                    Call appendResultObj(vaAllResults_sq, resultObj)
                Case OPT_SEQ_4bf
                    Call appendResultObj(vaAllResults_4bf, resultObj)
                Case OPT_SEQ_5bf
                    Call appendResultObj(vaAllResults_5bf, resultObj)
                Case OPT_SEQ_3mb
                    Call appendResultObj(vaAllResults_3mb, resultObj)
            End Select
        Next
    End Sub

            Sub appendResultObj(vaObj, resultObj)
                If Not checkOwnerExists(vaObj, resultObj) Then
                    vaObj.Append(resultObj)
                End If
            End Sub

            Function checkOwnerExists(vaObj, resultObj)
                If vaObj.Bound = -1 Then
                    checkOwnerExists = False
                    Exit Function
                End If

                Dim i, oResult
                For i = 0 To vaObj.Bound
                    Set oResult = vaObj.V(i)
                    If oResult.ResultOwner = resultObj.ResultOwner Then
                        If oResult.PostNum - resultObj.PostNum > 0 Then
                            Call vaObj.PopBySeq(i)
                            checkOwnerExists = False
                        Else
                            checkOwnerExists = True
                        End If
                        Exit Function
                    End If
                Next
                checkOwnerExists = False
            End Function

    Sub sortByAvgResult(vaObj, iSortType)
        If vaObj.Bound < 1 Then
            Exit Sub
        End If

        Dim i, j, oResult1, oResult2, iCompareResult
        If iSortType = SORT_BY_AVG Then
            For i = 0 To vaObj.Bound - 1
                Set oResult1 = vaObj.V(i)

                For j = i + 1 To vaObj.Bound
                    Set oResult2 = vaObj.V(j)
                    iCompareResult = oResult1.AvgResult - oResult2.AvgResult

                    If iCompareResult > 0 Then
                        Call vaObj.SwapTwoValues(i, j)
                        Set oResult1 = vaObj.V(i)
                    ElseIf iCompareResult = 0 Then
                        iCompareResult = compareAllResults(oResult1.SortedResults, oResult2.SortedResults)
                        If iCompareResult > 0 Then
                            Call vaObj.SwapTwoValues(i, j)
                            Set oResult1 = vaObj.V(i)
                        End If
                    End If
                Next
            Next
        ElseIf iSortType = SORT_BY_BEST Then
            For i = 0 To vaObj.Bound - 1
                Set oResult1 = vaObj.V(i)
                
                For j = i + 1 To vaObj.Bound
                    Set oResult2 = vaObj.V(j)
                    iCompareResult = oResult1.BestResult - oResult2.BestResult

                    If iCompareResult > 0 Then
                        Call vaObj.SwapTwoValues(i, j)
                        Set oResult1 = vaObj.V(i)
                    ElseIf iCompareResult = 0 Then
                        iCompareResult = compareAllResults(oResult1.SortedResults, oResult2.SortedResults)
                        If iCompareResult > 0 Then
                            Call vaObj.SwapTwoValues(i, j)
                            Set oResult1 = vaObj.V(i)
                        End If
                    End If
                Next
            Next
        ElseIf iSortType = SORT_BY_BEST_ANTI Then
            For i = 0 To vaObj.Bound - 1
                Set oResult1 = vaObj.V(i)
                
                For j = i + 1 To vaObj.Bound
                    Set oResult2 = vaObj.V(j)
                    iCompareResult = oResult1.BestResult - oResult2.BestResult

                    If iCompareResult < 0 Then
                        Call vaObj.SwapTwoValues(i, j)
                        Set oResult1 = vaObj.V(i)
                    ElseIf iCompareResult = 0 Then
                        iCompareResult = compareAllResults(oResult1.SortedResults, oResult2.SortedResults)
                        If iCompareResult < 0 Then
                            Call vaObj.SwapTwoValues(i, j)
                            Set oResult1 = vaObj.V(i)
                        End If
                    End If
                Next
            Next
        End If
    End Sub

            Function compareAllResults(sSortedResults1, sSortedResults2)
                Dim k, aTmp1, aTmp2, iDiff
                aTmp1 = Split(sSortedResults1)
                aTmp2 = Split(sSortedResults2)

                For k = 0 To UBound(aTmp1)
                    iDiff = aTmp1(k) - aTmp2(k)
                    If iDiff < 0 Then
                        compareAllResults = 0
                        Exit Function
                    ElseIf iDiff > 0 Then
                        compareAllResults = 1
                        Exit Function
                    End If
                Next
                
                compareAllResults = 0
            End Function

    Sub enableElementAfterSortAllResults()
        Call enableElement(ID_BUTTON_SAVE_ALL_RESULTS)
    End Sub

