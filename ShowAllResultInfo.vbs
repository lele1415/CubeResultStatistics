'*************************************************
'****show all result info.
'*************************************************
Dim mCrtShowSelectId, mCrtShowSeq

Sub showAllResultInfo()
    Call addValidResultInfoToList(vaAllValidResultInfo, ID_SELECT_VALID_RESULTS)
    Call addInvalidResultInfoToList(vaAllInvalidResultInfo, ID_SELECT_INVALID_RESULTS)
End Sub

        Sub addValidResultInfoToList(vaObj, selectId)
            Dim i, optionName
            For i = 0 To vaObj.Bound
                optionName = vaObj.V(i).PostNum & "楼 " & vaObj.V(i).ResultOwner
                Call addOption(selectId, optionName, i)
                If vaObj.V(i).IsBestBr Or vaObj.V(i).IsAvgBr Then
                    Call addOption(ID_SELECT_BR_RESULTS, optionName, i)
                End If
            Next
        End Sub

        Sub addInvalidResultInfoToList(vaObj, selectId)
            Dim i, optionName
            For i = 0 To vaObj.Bound
                optionName = vaObj.V(i).PostNum & "楼 " & vaObj.V(i).ResultOwner
                Call addOption(selectId, optionName, i)
            Next
        End Sub

Sub showResultInfo(selectId)
    Call clearAllSelectedOption(selectId)

    mCrtShowSelectId = selectId
    mCrtShowSeq = document.getElementById(selectId).value

    If mCrtShowSeq = "" Then Exit Sub

    Dim obj
    Select Case selectId
        Case ID_SELECT_VALID_RESULTS
            Set obj = vaAllValidResultInfo.V(mCrtShowSeq)
        Case ID_SELECT_INVALID_RESULTS
            Set obj = vaAllInvalidResultInfo.V(mCrtShowSeq)
        Case ID_SELECT_BR_RESULTS
            Set obj = vaAllValidResultInfo.V(mCrtShowSeq)
    End Select

    Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_POST_NUM, obj.PostNum)
    Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_POST_USER, obj.ResultOwner)
    Call setValueById(ID_TEXTAREA_SHOW_RESULT_OPT_NAME, obj.ResultOptSeq)

    If selectId <> ID_SELECT_INVALID_RESULTS Then
        Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_BEST_RESULT, obj.BestResult)
        Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_IS_BEST_BR, obj.IsBestBr)
        Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_AVG_RESULT, obj.AvgResult)
        Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_IS_AVG_BR, obj.IsAvgBr)
    Else
        Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_BEST_RESULT, "")
        Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_IS_BEST_BR, "")
        Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_AVG_RESULT, "")
        Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_IS_AVG_BR, "")
    End If

    Call setElementValue(ID_TEXTAREA_SHOW_RESULT_RESULT_TEXT, obj.ResultText)
    Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_PURE_RESULTS, obj.PureResults)
End Sub

Sub addNewResultInfo()
    Call clearAllSelectedOption("")
    Call clearAllResultInfoText()
    mCrtShowSelectId = ""
End Sub

Sub removeValidResult()
    If mCrtShowSelectId = ID_SELECT_VALID_RESULTS Then
        Dim oResult
        Set oResult = vaAllValidResultInfo.V(mCrtShowSeq)
        Call vaAllInvalidResultInfo.Append(oResult)
        Call addOption(ID_SELECT_INVALID_RESULTS, oResult.PostNum&"楼 "&oResult.ResultOwner, vaAllInvalidResultInfo.Bound)

        Call vaAllValidResultInfo.PopBySeq(mCrtShowSeq)
        Call removeOption(ID_SELECT_VALID_RESULTS, mCrtShowSeq)
        Call clearAllResultInfoText()
    End If
End Sub

Sub clearAllResultInfoText()
    Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_POST_NUM, "")
    Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_POST_USER, "")
    Call setValueById(ID_TEXTAREA_SHOW_RESULT_OPT_NAME, "")

    Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_BEST_RESULT, "")
    Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_IS_BEST_BR, "")
    Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_AVG_RESULT, "")
    Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_IS_AVG_BR, "")

    Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_RESULT_TEXT, "")
    Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_PURE_RESULTS, "")
End Sub

Sub submitNewResultInfo()
    Dim iNewPostNum, sNewPostUser, iNewOptSeq, sNewResultText, sNewPureResult, bIsValid

    iNewPostNum = getValueById(ID_TEXTAREA_SHOW_RESULT_POST_NUM)
    sNewPostUser = getValueById(ID_TEXTAREA_SHOW_RESULT_POST_USER)
    iNewOptSeq = getValueById(ID_TEXTAREA_SHOW_RESULT_OPT_NAME)
    sNewResultText = getValueById(ID_TEXTAREA_SHOW_RESULT_RESULT_TEXT)

    If iNewOptSeq = "" Then
        MsgBox("optName is empty!")
        Exit Sub
    End If

    Dim aTmp
    aTmp = getPureResult(sNewResultText, iNewOptSeq)
    sNewPureResult = aTmp(0)
    bIsValid = aTmp(1)

    If Not bIsValid Then MsgBox("is invalid result!") : Exit Sub

    Dim oVRI
    Select Case mCrtShowSelectId
        Case ID_SELECT_VALID_RESULTS
            Set oVRI = vaAllValidResultInfo.V(mCrtShowSeq)
            Call updateOldValidResultInfo(oVRI, iNewPostNum, sNewPostUser, iNewOptSeq, sNewResultText, sNewPureResult)
            Call checkAndAddBrOption(oVRI, iNewPostNum, sNewPostUser, mCrtShowSeq)
            Call updateNewResult(oVRI)

        Case ID_SELECT_INVALID_RESULTS
            Set oVRI = New ValidResultInfo
            Call createNewValidResultInfo(oVRI, iNewPostNum, sNewPostUser, iNewOptSeq, sNewResultText, sNewPureResult)
            Call addOption(ID_SELECT_VALID_RESULTS, iNewPostNum&"楼 "&sNewPostUser, vaAllValidResultInfo.Bound)
            Call removeOption(ID_SELECT_INVALID_RESULTS, mCrtShowSeq)
            Call checkAndAddBrOption(oVRI, iNewPostNum, sNewPostUser, vaAllValidResultInfo.Bound)

        Case ID_SELECT_BR_RESULTS
            Set oVRI = vaAllValidResultInfo.V(mCrtShowSeq)
            Call updateOldValidResultInfo(oVRI, iNewPostNum, sNewPostUser, iNewOptSeq, sNewResultText, sNewPureResult)
            Call checkAndRemoveBrOption(oVRI, mCrtShowSeq)
            Call updateNewResult(oVRI)

        Case ""
            Set oVRI = New ValidResultInfo
            Call createNewValidResultInfo(oVRI, iNewPostNum, sNewPostUser, iNewOptSeq, sNewResultText, sNewPureResult)
            Call addOption(ID_SELECT_VALID_RESULTS, iNewPostNum&"楼 "&sNewPostUser, vaAllValidResultInfo.Bound)
            Call checkAndAddBrOption(oVRI, iNewPostNum, sNewPostUser, vaAllValidResultInfo.Bound)
            Call updateNewResult(oVRI)
    End Select
End Sub

        Sub updateOldValidResultInfo(obj, iNewPostNum, sNewPostUser, iNewOptSeq, sNewResultText, sNewPureResult)
            obj.PostNum = iNewPostNum
            obj.ResultOwner = sNewPostUser
            obj.ResultOptSeq = iNewOptSeq
            obj.ResultText = sNewResultText
            obj.PureResults = sNewPureResult

            Call getSortedResult(sNewPureResult, iNewOptSeq, obj)
        End Sub

        Sub createNewValidResultInfo(obj, iNewPostNum, sNewPostUser, iNewOptSeq, sNewResultText, sNewPureResult)
            obj.PostNum = iNewPostNum
            obj.ResultOwner = sNewPostUser
            obj.ResultOptSeq = iNewOptSeq
            obj.ResultText = sNewResultText
            obj.PureResults = sNewPureResult

            Call getSortedResult(sNewPureResult, iNewOptSeq, obj)

            Call vaAllValidResultInfo.Append(obj)
            'Call vaAllInvalidResultInfo.PopBySeq(mCrtShowSeq)
        End Sub

        Sub checkAndAddBrOption(obj, iNewPostNum, sNewPostUser, optionValue)
            If obj.IsBestBr Or obj.IsAvgBr Then
                Call addOption(ID_SELECT_BR_RESULTS, iNewPostNum&"楼 "&sNewPostUser, optionValue)
            End If
        End Sub

        Sub checkAndRemoveBrOption(obj, optionValue)
            If Not(obj.IsBestBr Or obj.IsAvgBr) Then
                Call removeOption(ID_SELECT_BR_RESULTS, optionValue)
            End If
        End Sub

        Sub updateNewResult(obj)
            Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_BEST_RESULT, obj.BestResult)
            Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_IS_BEST_BR, obj.IsBestBr)
            Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_AVG_RESULT, obj.AvgResult)
            Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_IS_AVG_BR, obj.IsAvgBr)
            Call setInnerHtmlById(ID_TEXTAREA_SHOW_RESULT_PURE_RESULTS, obj.PureResults)
        End Sub

Sub selectNextInvalidResult()
    Dim mCrtShowInvalidSeq
    mCrtShowInvalidSeq = document.getElementById(ID_SELECT_INVALID_RESULTS).value
    If mCrtShowInvalidSeq = "" Then Exit Sub
    Call selectAnotherOption(ID_SELECT_INVALID_RESULTS, mCrtShowInvalidSeq, 0)
End Sub

Sub selectPrevInvalidResult()
    Dim mCrtShowInvalidSeq
    mCrtShowInvalidSeq = document.getElementById(ID_SELECT_INVALID_RESULTS).value
    If mCrtShowInvalidSeq = "" Then Exit Sub
    Call selectAnotherOption(ID_SELECT_INVALID_RESULTS, mCrtShowInvalidSeq, 1)
End Sub
