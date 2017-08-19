'*************************************************
'****get result.
'*************************************************
Function getPureResult(sMessageText, sOptName)
    Dim aPickedResult, iNeedNum, aPureResult, isValid
    aPickedResult = pickMachedResults(sMessageText, sOptName)
    If isArray(aPickedResult) Then
        iNeedNum = getNeedNumOfResult(sOptName)
        isValid = checkResultsNum(aPickedResult, iNeedNum)
        getPureResult = Array(joinArrayWithSpace(aPickedResult), isValid)
    Else
        getPureResult = Array("", False)
    End If
End Function

        Function getNeedNumOfResult(sOptName)
            If (sOptName = "333bf") Or _
                    (sOptName = "333fm") Or _
                    (sOptName = "666") Or _
                    (sOptName = "777") Then
                getNeedNumOfResult = 3
            Else
                getNeedNumOfResult = 5
            End If
        End Function

        Function pickMachedResults(sMessageText, sOptName)
            Dim aTmpResult, sTmp, sTmpResult
            'MsgBox("sMessageText="&sMessageText&Vblf&_
                    '"replaceCharacterInResultStr(sMessageText)="&replaceCharacterInResultStr(sMessageText))
            aTmpResult = Split(replaceCharacterInResultStr(cutResultStrByKeyword(sMessageText)))
            For i = 0 To safeUBound(aTmpResult, "pickMachedResults", 1)
                If sOptName = "333fm" Then
                    sTmp = formatResultStrForFm(aTmpResult(i))
                Else
                    sTmp = formatResultStr(aTmpResult(i))
                End If

                If IsNumeric(sTmp) Then
                    sTmpResult = sTmpResult & " " & sTmp
                End If
            Next
            'MsgBox(sTmpResult)

            pickMachedResults = Split(Trim(sTmpResult))
        End Function

        Function cutResultStrByKeyword(sMessageText)
            Dim iKeywordInStr, aKeyword, sTmp
            aKeyword = Array("成绩列表","详细时间")
            sTmp = sMessageText
            For i = 0 To safeUBound(aKeyword, "cutResultStrByKeyword", 1)
                iKeywordInStr = InStr(sMessageText, aKeyword(i))
                If iKeywordInStr > 0 Then
                    sTmp = safeMid(sMessageText, iKeywordInStr+4, 0, "cutResultStrByKeyword")
                    Exit For
                End If
            Next
            cutResultStrByKeyword = sTmp
        End Function

        Function checkResultsNum(aPickedResult, iNeedNum)
            Dim isValid
            If safeUBound(aPickedResult, "checkResultsNum", 1) + 1 = iNeedNum Then
                isValid = True
            Else
                isValid = False
            End If
            checkResultsNum = isValid
        End Function