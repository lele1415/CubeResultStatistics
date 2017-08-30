'*************************************************
'****get result.
'*************************************************
Dim aResultHeadword : aResultHeadword = Array("成绩列表","详细时间")

Function getPureResult(sResultText, iOptSeq)
    Dim svaPickedResult, isValid, iNeedNum
    iNeedNum = vaOptInfo.V(iOptSeq).NeedNum
    Set svaPickedResult = pickMachedResults(sResultText, iOptSeq)
    If svaPickedResult.Bound = -1 Then getPureResult = "" : Exit Function

    isValid = checkResultsNum(svaPickedResult.Bound, iNeedNum)
    getPureResult = Array(svaPickedResult.ToString(), isValid)
End Function

        Function pickMachedResults(sResultText, iOptSeq)
            Dim aTmpResult, sTmp
            'MsgBox("sResultText="&sResultText&Vblf&_
                    '"replaceCharacterInResultStr(sResultText)="&replaceCharacterInResultStr(sResultText))
            aTmpResult = Split(replaceCharacterInResultStr(cutResultStrByKeyword(sResultText)))
            If Not isArray(aTmpResult) Then pickMachedResults = "" : Exit Function

            Dim i, svaObj
            Set svaObj = New SimpleVariableArray

            For i = 0 To UBound(aTmpResult)
                sTmp = formatResultStr(aTmpResult(i), iOptSeq)
                
                If IsNumeric(sTmp) Then
                    svaObj.Append(sTmp)
                End If
            Next
            'MsgBox(sTmpResult)

            Set pickMachedResults = svaObj
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

        Function checkResultsNum(iBound, iNeedNum)
            Dim isValid
            If iBound + 1 = iNeedNum Then
                isValid = True
            Else
                isValid = False
            End If
            checkResultsNum = isValid
        End Function