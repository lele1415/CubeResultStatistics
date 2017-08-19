'*************************************************
'****get result.
'*************************************************
Dim aResultHeadword : aResultHeadword = Array("成绩列表","详细时间")

Function getPureResult(sResultText, iOptSeq)
    Dim aPickedResult, aPureResult, isValid, iNeedNum
    iNeedNum = vaOptInfo.V(iOptSeq).NeedNum
    aPickedResult = pickMachedResults(sResultText, iOptSeq, iNeedNum)
    If Not isArray(aPickedResult) Then getPureResult = "" : Exit Function

    isValid = checkResultsNum(aPickedResult, iNeedNum)
    getPureResult = Array(joinArrayWithSpace(aPickedResult), isValid)
End Function

        Function pickMachedResults(sResultText, iOptSeq, iNeedNum)
            Dim aTmpResult, sTmp, sTmpResult
            'MsgBox("sResultText="&sResultText&Vblf&_
                    '"replaceCharacterInResultStr(sResultText)="&replaceCharacterInResultStr(sResultText))
            aTmpResult = Split(replaceCharacterInResultStr(cutResultStrByKeyword(sResultText)))
            If Not isArray(aTmpResult) Then pickMachedResults = "" : Exit Function

            Dim i, iMached
            iMached = 0
            For i = 0 To UBound(aTmpResult)
                sTmp = formatResultStr(aTmpResult(i), iOptSeq)
                
                If IsNumeric(sTmp) Then
                    sTmpResult = sTmpResult & " " & sTmp
                    iMached = iMached + 1
                End If

                If iMached = iNeedNum Then Exit For
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