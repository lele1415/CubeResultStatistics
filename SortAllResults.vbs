'*************************************************
'****sort all results.
'*************************************************
Function sortAllResults()
    sortByOptName()

    sortByAvgResult aTmpUserText_333, aTmpPureResults_333, aTmpSortedResults_333, 0
    sortByAvgResult aTmpUserText_444, aTmpPureResults_444, aTmpSortedResults_444, 0
    sortByAvgResult aTmpUserText_555, aTmpPureResults_555, aTmpSortedResults_555, 0
    sortByAvgResult aTmpUserText_222, aTmpPureResults_222, aTmpSortedResults_222, 0
    sortByAvgResult aTmpUserText_333bf, aTmpPureResults_333bf, aTmpSortedResults_333bf, 1
    sortByAvgResult aTmpUserText_333oh, aTmpPureResults_333oh, aTmpSortedResults_333oh, 0
    sortByAvgResult aTmpUserText_333fm, aTmpPureResults_333fm, aTmpSortedResults_333fm, 0
    sortByAvgResult aTmpUserText_py, aTmpPureResults_py, aTmpSortedResults_py, 0
    sortByAvgResult aTmpUserText_mega, aTmpPureResults_mega, aTmpSortedResults_mega, 0
    sortByAvgResult aTmpUserText_sq, aTmpPureResults_sq, aTmpSortedResults_sq, 0
    sortByAvgResult aTmpUserText_clock, aTmpPureResults_clock, aTmpSortedResults_clock, 0
    sortByAvgResult aTmpUserText_sk, aTmpPureResults_sk, aTmpSortedResults_sk, 0
    sortByAvgResult aTmpUserText_666, aTmpPureResults_666, aTmpSortedResults_666, 0
    sortByAvgResult aTmpUserText_777, aTmpPureResults_777, aTmpSortedResults_777, 0

    MsgBox("OK")
End Function

    Function sortByOptName()
        Dim aAllValidOptName, count, aTmpUserText(), aTmpPureResults(), aTmpSortedResults()
        aAllValidOptName = getAllOptName()
        count = 0
        

        For i = 0 To UBound(aAllValidOptName)
            ReDim Preserve aTmpUserText(0)
            ReDim Preserve aTmpPureResults(0)
            ReDim Preserve aTmpSortedResults(0)
            aTmpUserText(0) = ""
            aTmpPureResults(0) = ""
            aTmpSortedResults(0) = ""
            count = 0

            For j = 0 To iTmpResultCount - 1
                If aAllIsValid(j) And _
                        aAllOptName(j) = aAllValidOptName(i) And _
                        getExistInArray(aTmpUserText, aAllUserText(j)) = "" Then
                    ReDim Preserve aTmpUserText(count)
                    ReDim Preserve aTmpPureResults(count)
                    ReDim Preserve aTmpSortedResults(count)

                    aTmpUserText(count) = aAllUserText(j)
                    aTmpPureResults(count) = aAllPureResults(j)
                    aTmpSortedResults(count) = aAllSortedResults(j)
                    count = count + 1
                End If
            Next

            Select Case aAllValidOptName(i)
                Case "333"
                    aTmpUserText_333 = aTmpUserText
                    aTmpPureResults_333 = aTmpPureResults
                    aTmpSortedResults_333 = aTmpSortedResults
                Case "444"
                    aTmpUserText_444 = aTmpUserText
                    aTmpPureResults_444 = aTmpPureResults
                    aTmpSortedResults_444 = aTmpSortedResults
                Case "555"
                    aTmpUserText_555 = aTmpUserText
                    aTmpPureResults_555 = aTmpPureResults
                    aTmpSortedResults_555 = aTmpSortedResults
                Case "222"
                    aTmpUserText_222 = aTmpUserText
                    aTmpPureResults_222 = aTmpPureResults
                    aTmpSortedResults_222 = aTmpSortedResults
                Case "333bf"
                    aTmpUserText_333bf = aTmpUserText
                    aTmpPureResults_333bf = aTmpPureResults
                    aTmpSortedResults_333bf = aTmpSortedResults
                Case "333oh"
                    aTmpUserText_333oh = aTmpUserText
                    aTmpPureResults_333oh = aTmpPureResults
                    aTmpSortedResults_333oh = aTmpSortedResults
                Case "333fm"
                    aTmpUserText_333fm = aTmpUserText
                    aTmpPureResults_333fm = aTmpPureResults
                    aTmpSortedResults_333fm = aTmpSortedResults
                Case "py"
                    aTmpUserText_py = aTmpUserText
                    aTmpPureResults_py = aTmpPureResults
                    aTmpSortedResults_py = aTmpSortedResults
                Case "mega"
                    aTmpUserText_mega = aTmpUserText
                    aTmpPureResults_mega = aTmpPureResults
                    aTmpSortedResults_mega = aTmpSortedResults
                Case "sq"
                    aTmpUserText_sq = aTmpUserText
                    aTmpPureResults_sq = aTmpPureResults
                    aTmpSortedResults_sq = aTmpSortedResults
                Case "clock"
                    aTmpUserText_clock = aTmpUserText
                    aTmpPureResults_clock = aTmpPureResults
                    aTmpSortedResults_clock = aTmpSortedResults
                Case "sk"
                    aTmpUserText_sk = aTmpUserText
                    aTmpPureResults_sk = aTmpPureResults
                    aTmpSortedResults_sk = aTmpSortedResults
                Case "666"
                    aTmpUserText_666 = aTmpUserText
                    aTmpPureResults_666 = aTmpPureResults
                    aTmpSortedResults_666 = aTmpSortedResults
                Case "777"
                    aTmpUserText_777 = aTmpUserText
                    aTmpPureResults_777 = aTmpPureResults
                    aTmpSortedResults_777 = aTmpSortedResults
            End Select
        Next
    End Function

    Function sortByAvgResult(aTmpUserText, aTmpPureResults, aTmpSortedResults, iSeq)
        Dim lenTmpUserText, flag
        lenTmpUserText = safeUBound(aTmpUserText, "sortByAvgResult 111", 1)
        flag = 0

        If lenTmpUserText > 0 Then
            For i = 0 To lenTmpUserText - 1
                For j = i + 1 To lenTmpUserText
                    Dim aResults1, aResults2
                    aResults1 = Split(aTmpSortedResults(i))
                    aResults2 = Split(aTmpSortedResults(j))

                    If iSeq = 0 Then
                        flag = compareTwoResult(aResults1(iSeq), aResults2(iSeq), i, j, aTmpUserText, aTmpPureResults, aTmpSortedResults)
                    End If

                    If (iSeq = 1 Or flag = 1) Then
                        For k = 1 To safeUBound(aResults1, "sortByAvgResult 222", 1) - 2
                            flag = compareTwoResult(aResults1(k), aResults2(k), i, j, aTmpUserText, aTmpPureResults, aTmpSortedResults)
                            If flag = 0 Then
                                Exit For
                            End If
                        Next
                    End If
                Next
            Next
        End If
    End Function

        Function compareTwoResult(iResult1, iResult2, i, j, aTmpUserText, aTmpPureResults, aTmpSortedResults)
            Dim flag
            flag = 0
            'If Not IsNumeric(iResult1) Then
            '    MsgBox("iResult1=" & iResult1 & Vblf &_
            '            "aTmpUserText(i)=" & aTmpUserText(i) & Vblf &_
            '            "aTmpPureResults(i)=" & aTmpPureResults(i) & Vblf &_
            '            "aTmpSortedResults(i)=" & aTmpSortedResults(i))
            'End If
            If iResult1 - iResult2 > 0 Then
                swapTwoStrings aTmpUserText(i), aTmpUserText(j)
                swapTwoStrings aTmpPureResults(i), aTmpPureResults(j)
                swapTwoStrings aTmpSortedResults(i), aTmpSortedResults(j)
            ElseIf iResult1 - iResult2 = 0 Then
                flag = 1
            End If

            compareTwoResult = flag
        End Function
