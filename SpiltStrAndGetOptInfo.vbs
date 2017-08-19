'*************************************************
'****split str and get Opt info.
'*************************************************
Function spiltStrAndGetOptInfo()
    If Not testFlag Then
        initTxtFile(uValidResultTxt)
        initTxtFile(uInvalidResultTxt)
        loadAllRecords()
    End If

    Dim aOptInfo()
    ReDim aOptInfo(1,0)
    For i = 0 To iOriginDataArrayLength
        If checkPostNumIsNeed(aPostNum(i)) And checkUserIsNeed(aUserText(i)) Then
            setInnerHtmlById "postNum_id", aPostNum(i)
            
            Dim iOptCount, aSplitMessageText, bGetOptName
            iOptCount = 0
            'If aPostNum(i) = 59 Then
                'MsgBox(replaceHtmlStr(aMessageText(i)))
            'End If
            aSplitMessageText = Split(replaceHtmlStr(aMessageText(i)))
            bGetOptName = False

            getOptInfoArray aOptInfo, iOptCount, aSplitMessageText, bGetOptName

            If bGetOptName Then
                For j = 0 To safeUBound(aOptInfo, "spiltStrAndGetOptInfo", 2) - 1
                    'MsgBox( "aPostNum(i)="&aPostNum(i)&Vblf&_
                            '"aOptInfo(1, j)="&aOptInfo(1, j)&Vblf&_
                            '"aOptInfo(1, j+1)="&aOptInfo(1, j+1) )
                    If aOptInfo(1, j + 1) - aOptInfo(1, j) > 1 Then
                        Dim aCutedUserResult, aTmp, sPureResult, isValid, sMessageText, sUserText, iPostNum, sOptName
                        aCutedUserResult = cutArray(aSplitMessageText, aOptInfo(1, j) + 1, aOptInfo(1, j + 1) - 1)
                        
                        sMessageText = joinArrayWithSpace(aCutedUserResult)
                        sUserText = aUserText(i)
                        iPostNum = aPostNum(i)
                        sOptName = aOptInfo(0, j)

                        aTmp = getPureResult(sMessageText, sUserText, iPostNum, sOptName)
                        sPureResult = aTmp(0)
                        isValid = aTmp(1)
                        saveResultToArray sMessageText, sPureResult, sUserText, iPostNum, sOptName, isValid
                    Else
                        saveResultToArray joinArrayWithSpace(aSplitMessageText), "", aUserText(i), aPostNum(i), aOptInfo(0, j), False
                    End If
                Next
            Else
                saveResultToArray joinArrayWithSpace(aSplitMessageText), "", aUserText(i), aPostNum(i), "unknow", False
            End If

            If testFlag Then
                Exit For
            End If
        End If
    Next

    getResultDoneFlag = True
    MsgBox("OK")
End Function

        Function getOptInfoArray(aOptInfo, iOptCount, aSplitMessageText, bGetOptName)
            Dim sOptName, lenSplitUserResult
            lenSplitUserResult = safeUBound(aSplitMessageText, "getOptInfoArray", 1)

            For i = 0 To lenSplitUserResult
                sOptName = checkOptNameStr(replaceStrForOptName(aSplitMessageText(i)))
                If sOptName <> "" Then
                    ReDim Preserve aOptInfo(1, iOptCount)
                    aOptInfo(0, iOptCount) = sOptName
                    aOptInfo(1, iOptCount) = i
                    iOptCount = iOptCount + 1
                    bGetOptName = True
                End If
            Next

            ReDim Preserve aOptInfo(1, iOptCount)
            aOptInfo(0, iOptCount) = ""
            aOptInfo(1, iOptCount) = lenSplitUserResult + 1
        End Function

        Function checkOptNameStr(str)
            Dim aAllOptName, bExit
            aAllOptName = getAllOptName()
            bExit = False

            For i = 0 To safeUBound(aAllOptName, "checkOptNameStr 111", 1)
                Dim aOptNameForCheck
                aOptNameForCheck = getOptCheckStrArray(aAllOptName(i))

                For j = 0 To safeUBound(aOptNameForCheck, "checkOptNameStr 222", 1)
                    'MsgBox("aOptNameForCheck(j)="&aOptNameForCheck(j)&" str="&str)
                    If str = aOptNameForCheck(j) Then
                        checkOptNameStr = aAllOptName(i)
                        bExit = True
                        Exit For
                    End If
                Next

                If bExit Then
                    Exit For
                End If
            Next

            If Not bExit Then
                checkOptNameStr = ""
            End If
        End Function

        Function checkPostNumIsNeed(iPostNum)
            Dim exceptPostNumList, flag
            exceptPostNumList = Split(getValueById("exceptPostNum_id"), " ")
            flag = True
            For i = 0 To safeUBound(exceptPostNumList, "checkPostNumIsNeed", 1)
                If iPostNum = exceptPostNumList(i) Then
                    flag = False
                    Exit For
                End If
            Next
            checkPostNumIsNeed = flag
        End Function

        Function checkUserIsNeed(name)
            Dim exceptNameList, flag
            exceptNameList = Split(getValueById("except_name_id"), " ")
            flag = True
            For i = 0 To safeUBound(exceptNameList, "checkUserIsNeed", 1)
                If name = exceptNameList(i) Then
                    flag = False
                    Exit For
                End If
            Next
            checkUserIsNeed = flag
        End Function
        