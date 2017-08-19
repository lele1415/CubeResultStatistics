'*************************************************
'****save result to array.
'*************************************************
Function saveResultToArray(sMessageText, sPureResult, sUserText, iPostNum, sOptName, isValid)
    Dim slistId, sTxtPath
    If isValid Then
        slistId = "validResultList_id"
        sTxtPath = uValidResultTxt
    Else
        slistId = "invalidResultList_id"
        sTxtPath = uInvalidResultTxt
    End If

    ReDim Preserve aAllPostName(iTmpResultCount)
    ReDim Preserve aAllUserText(iTmpResultCount)
    ReDim Preserve aAllOptName(iTmpResultCount)
    ReDim Preserve aAllMessageText(iTmpResultCount)
    ReDim Preserve aAllPureResults(iTmpResultCount)
    ReDim Preserve aAllIsValid(iTmpResultCount)
    ReDim Preserve aAllSortedResults(iTmpResultCount)
    aAllPostName(iTmpResultCount) = iPostNum
    aAllUserText(iTmpResultCount) = sUserText
    aAllOptName(iTmpResultCount) = sOptName
    aAllMessageText(iTmpResultCount) = sMessageText
    aAllPureResults(iTmpResultCount) = sPureResult
    aAllIsValid(iTmpResultCount) = isValid
    aAllSortedResults(iTmpResultCount) = getAvgAndSortedResults(sOptName, sPureResult, isValid)

    ReDim Preserve aAllSortedResults(iTmpResultCount)

    addOption slistId, iPostNum&"楼 "&sUserText, iTmpResultCount

    If InStr(aAllSortedResults(iTmpResultCount), "True") Then
        addOption "brResultList_id", iPostNum&"楼 "&sUserText, iTmpResultCount
    End If

    iTmpResultCount = iTmpResultCount + 1

    Dim oTxt
    Set oTxt = Fso.OpenTextFile(sTxtPath, 8, False, True)

    oTxt.WriteLine(iPostNum)
    oTxt.WriteLine(sUserText)
    oTxt.WriteLine(sMessageText)
    oTxt.WriteLine(sOptName)
    oTxt.WriteLine(sPureResult)
    oTxt.WriteLine()

    oTxt.Close
    Set oTxt = Nothing
End Function

Function showResult(slistId)
    seqOfShowingResult = document.getElementById(slistId).value
    listOfShowingResult = slistId

    setInnerHtmlById "showResult_post_id", aAllPostName(seqOfShowingResult)
    setInnerHtmlById "showResult_user_id", aAllUserText(seqOfShowingResult)
    setValueById "showResult_opt_id", aAllOptName(seqOfShowingResult)
    setInnerHtmlById "showResult_msg_id", aAllMessageText(seqOfShowingResult)
    setInnerHtmlById "showResult_rst_id", aAllPureResults(seqOfShowingResult)
End Function

Function submitNewMessageText()
    Dim iPostNum, sUserText, sOptName, sNewMessageText, aTmp, sNewPureResult, isValid

    iPostNum = aAllPostName(seqOfShowingResult)
    sUserText = aAllUserText(seqOfShowingResult)
    sOptName = document.getElementById("showResult_opt_id").value
    sNewMessageText = document.getElementById("showResult_msg_id").value

    aTmp = getPureResult(sNewMessageText, sUserText, iPostNum, sOptName)
    sNewPureResult = aTmp(0)
    isValid = aTmp(1)

    If isValid Then
        aAllPureResults(seqOfShowingResult) = sNewPureResult
        aAllOptName(seqOfShowingResult) = sOptName
        aAllIsValid(seqOfShowingResult) = isValid
        aAllSortedResults(seqOfShowingResult) = getAvgAndSortedResults(sOptName, sNewPureResult, isValid)

        If listOfShowingResult = "invalidResultList_id" Then
            addOption "validResultList_id", iPostNum&"楼 "&sUserText, seqOfShowingResult
            removeOption "invalidResultList_id", seqOfShowingResult
        ELseIf listOfShowingResult = "brResultList_id" Then
            removeOption "brResultList_id", seqOfShowingResult
        End If

        If InStr(aAllSortedResults(seqOfShowingResult), "True") Then
            addOption "brResultList_id", iPostNum&"楼 "&sUserText, seqOfShowingResult
        End If
    Else
        MsgBox("is invalid result!")
    End If
End Function
