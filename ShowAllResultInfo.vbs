'*************************************************
'****show all result info.
'*************************************************
Const ID_SELECT_VALID_RESULT = "validResultList_id"
Const ID_SELECT_INVALID_RESULT = "invalidResultList_id"
Const ID_TEXT_POST_NUM = "showResult_post_num"
Const ID_TEXT_POST_USER = "showResult_post_user"
Const ID_TEXT_OPT_NAME = "showResult_opt_name"
Const ID_TEXT_RESULT_TEXT = "showResult_result_text"
Const ID_TEXT_PURE_RESULTS = "showResult_pure_results"

Dim mCrtShowSelectId, mCrtShowSeq

Sub showAllResultInfo()
    Call addResultInfoToList(vaAllValidResultInfo, ID_SELECT_VALID_RESULT)
    Call addResultInfoToList(vaAllInvalidResultInfo, ID_SELECT_INVALID_RESULT)
End Sub

        Sub addResultInfoToList(vaObj, selectId)
            Dim i, optionName
            For i = 0 To vaObj.Bound
                optionName = vaObj.V(i).PostNum & "楼 " & vaObj.V(i).ResultOwner
                Call addOption(selectId, optionName, i)
            Next
        End Sub

Sub showResultInfo(selectId)
    mCrtShowSelectId = selectId
    mCrtShowSeq = document.getElementById(selectId).value

    Dim obj
    Select Case selectId
        Case ID_SELECT_VALID_RESULT
            Set obj = vaAllValidResultInfo.V(mCrtShowSeq)
        Case ID_SELECT_INVALID_RESULT
            Set obj = vaAllInvalidResultInfo.V(mCrtShowSeq)
    End Select

    Call setInnerHtmlById(ID_TEXT_POST_NUM, obj.PostNum)
    Call setInnerHtmlById(ID_TEXT_POST_USER, obj.ResultOwner)
    Call setValueById(ID_TEXT_OPT_NAME, obj.OptSeq)
    Call setInnerHtmlById(ID_TEXT_RESULT_TEXT, obj.ResultText)
    Call setInnerHtmlById(ID_TEXT_PURE_RESULTS, obj.PureResults)
End Sub

Sub submitNewResultInfo()
    Dim newPostNum, newPostUser, newOptSeq, newResultText, aTmp, sNewPureResult, isValid

    newPostNum = getValueById(ID_TEXT_POST_NUM)
    newPostUser = getValueById(ID_TEXT_POST_USER)
    newOptSeq = getValueById(ID_TEXT_OPT_NAME)
    newResultText = getValueById(ID_TEXT_RESULT_TEXT)

    aTmp = getPureResult(newResultText, newOptSeq)

    If Not aTmp(1) Then MsgBox("is invalid result!") : Exit Sub

    Dim oNew : Set oNew = New ValidResultInfo
    oNew.PostNum = newPostNum
    oNew.ResultOwner = newPostUser
    oNew.ResultOpt = vaOptInfo.V(newOptSeq)
    oNew.ResultText = newResultText
    oNew.PureResults = aTmp(0)
    Call vaAllValidResultInfo.Append(oNew)
    Call vaAllInvalidResultInfo.PopBySeq(mCrtShowSeq)

    If mCrtShowSelectId = ID_SELECT_INVALID_RESULT Then
        addOption ID_SELECT_VALID_RESULT, newPostNum&"楼 "&vaOptInfo.V(newOptSeq).FullName, vaAllValidResultInfo.Bound
        removeOption ID_SELECT_INVALID_RESULT, mCrtShowSeq
    ELseIf listOfShowingResult = "brResultList_id" Then
        removeOption "brResultList_id", seqOfShowingResult
    End If

    If InStr(aAllSortedResults(seqOfShowingResult), "True") Then
        addOption "brResultList_id", iPostNum&"楼 "&sUserText, seqOfShowingResult
    End If
End Sub
