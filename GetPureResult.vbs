'*************************************************
'****get result.
'*************************************************
Dim aResultHeadword : aResultHeadword = Array("成绩列表","详细时间")

Function getPureResult(sResultText, iOptSeq)
    Dim vaPickedResult, isValid, iNeedNum
    iNeedNum = vaOptInfo.V(iOptSeq).NeedNum
    Set vaPickedResult = pickMatchedResults(sResultText, iOptSeq)
    If vaPickedResult.Bound = -1 Then getPureResult = Array("", False) : Exit Function

    isValid = checkResultsNum(vaPickedResult.Bound, iNeedNum)
    getPureResult = Array(vaPickedResult.ToStringWithSpace(), isValid)
End Function

        Function pickMatchedResults(sResultText, iOptSeq)
            Dim aTmpResult, sTmp
            'MsgBox("sResultText="&sResultText&Vblf&_
                    '"replaceCharacterInResultStr(sResultText)="&replaceCharacterInResultStr(sResultText))
            aTmpResult = Split(replaceCharacterInResultStr(cutResultStrByKeyword(sResultText)))
            If Not isArray(aTmpResult) Then pickMatchedResults = "" : Exit Function

            Dim i, vaObj
            Set vaObj = New VariableArray
            Call vaObj.SetPreBound(4)

            For i = 0 To UBound(aTmpResult)
                Call formatResultStr(aTmpResult(i), iOptSeq, vaObj)
            Next
            'MsgBox(sTmpResult)

            Set pickMatchedResults = vaObj
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