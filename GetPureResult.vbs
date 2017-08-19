'*************************************************
'****get result.
'*************************************************
Dim aResultHeadword : aResultHeadword = Array("成绩列表","详细时间")

Function getPureResult(sResultText, iOptSeq)
    Dim aPickedResult, aPureResult, isValid
    aPickedResult = pickMachedResults(sResultText, iOptSeq)
    If Not isArray(aPickedResult) Then getPureResult = "" : Exit Function

    isValid = checkResultsNum(aPickedResult, vaOptInfo.V(iOptSeq).NeedNum)
    getPureResult = Array(joinArrayWithSpace(aPickedResult), isValid)
End Function

        Function pickMachedResults(sResultText, iOptSeq)
            Dim aTmpResult, sTmp, sTmpResult
            'MsgBox("sResultText="&sResultText&Vblf&_
                    '"replaceCharacterInResultStr(sResultText)="&replaceCharacterInResultStr(sResultText))
            aTmpResult = Split(replaceCharacterInResultStr(cutResultStrByKeyword(sResultText)))
            If Not isArray(aTmpResult) Then pickMachedResults = "" : Exit Function

            Dim formatFun
            If iOptSeq = OPT_SEQ_3fm Then
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

        Function cutResultStrByKeyword(sResultText)
            Dim iKeywordInStr, sTmp, sKeyword
            sTmp = sResultText
            For i = 0 To UBound(aResultHeadword)
                sKeyword = aResultHeadword(i)
                iKeywordInStr = InStr(sResultText, sKeyword)
                If iKeywordInStr > 0 Then
                    sTmp = safeMid(sResultText, iKeywordInStr + Len(sKeyword), 0, "cutResultStrByKeyword")
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