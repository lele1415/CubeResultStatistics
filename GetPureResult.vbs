'*************************************************
'****get result.
'*************************************************
Dim aResultHeadword : aResultHeadword = Array("成绩列表","详细时间")

Function getPureResult(sMessageText, sOptName)
    Dim aPickedResult, iNeedNum, aPureResult, isValid
    aPickedResult = pickMachedResults(sMessageText, sOptName)
    If Not isArray(aPickedResult) Then getPureResult = "" : Exit Function

    iNeedNum = getNeedNumOfResult(sOptName)
    isValid = checkResultsNum(aPickedResult, iNeedNum)
    getPureResult = Array(joinArrayWithSpace(aPickedResult), isValid)
End Function

        Function getNeedNumOfResult(sOptName)
            If (sOptName = OPT_3bf) Or _
                    (sOptName = OPT_3fm) Or _
                    (sOptName = OPT_666) Or _
                    (sOptName = OPT_777) Then
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
            If Not isArray(aTmpResult) Then pickMachedResults = "" : Exit Function

            Dim formatFun
            If sOptName = "333fm" Then
                formatFun = "formatResultStrForFm"
            Else
                formatFun = "formatResultStr"
            End If

            For i = 0 To UBound(aTmpResult)
                Execute "sTmp = " & formatFun & "(aTmpResult(i))"
                If IsNumeric(sTmp) Then
                    sTmpResult = sTmpResult & " " & sTmp
                End If
            Next
            'MsgBox(sTmpResult)

            pickMachedResults = Split(Trim(sTmpResult))
        End Function

        Function cutResultStrByKeyword(sMessageText)
            Dim iKeywordInStr, sTmp, sKeyword
            sTmp = sMessageText
            For i = 0 To UBound(aResultHeadword)
                sKeyword = aResultHeadword(i)
                iKeywordInStr = InStr(sMessageText, sKeyword)
                If iKeywordInStr > 0 Then
                    sTmp = safeMid(sMessageText, iKeywordInStr + Len(sKeyword), 0, "cutResultStrByKeyword")
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