'*************************************************
'****split str and get Opt info.
'*************************************************
Dim vaAllValidResultInfo : Set vaAllValidResultInfo = New VariableArray
Dim vaAllInvalidResultInfo : Set vaAllInvalidResultInfo = New VariableArray
Dim vaAllOptLocationInfo : Set vaAllOptLocationInfo = New VariableArray

Const ID_CRT_POST_NUM = "crt_post_num"

Const ID_SELECT_VALID_RESULTS = "valid_results"
Const ID_SELECT_INVALID_RESULTS = "invalid_results"
Const ID_SELECT_BR_RESULTS = "br_results"
Const ID_BUTTON_ADD_NEW_RESULT_INFO = "add_new_result_info"
Const ID_BUTTON_REMOVE_VALID_RESULT = "remove_valid_result"
Const ID_TEXTAREA_SHOW_RESULT_POST_NUM = "show_result_post_num"
Const ID_TEXTAREA_SHOW_RESULT_POST_USER = "show_result_post_user"
Const ID_TEXTAREA_SHOW_RESULT_OPT_NAME = "show_result_opt_name"
Const ID_TEXTAREA_SHOW_RESULT_BEST_RESULT = "show_result_best_result"
Const ID_TEXTAREA_SHOW_RESULT_IS_BEST_BR = "show_result_is_best_br"
Const ID_TEXTAREA_SHOW_RESULT_AVG_RESULT = "show_result_avg_result"
Const ID_TEXTAREA_SHOW_RESULT_IS_AVG_BR = "show_result_is_avg_br"
Const ID_TEXTAREA_SHOW_RESULT_RESULT_TEXT = "show_result_result_text"
Const ID_TEXTAREA_SHOW_RESULT_PURE_RESULTS = "show_result_pure_results"
Const ID_BUTTON_SUBMIT_NEW_RESULT_INFO = "submit_new_result_info"
Const ID_BUTTON_SELECT_NEXT_INVALID_RESULT = "select_next_invalid_result"
Const ID_BUTTON_SELECT_PREV_INVALID_RESULT = "select_prev_invalid_result"
Const ID_BUTTON_SORT_ALL_RESULTS = "sort_all_results"

Dim oCrtPostInfo, iCrtPostInfoSeq, iMaxPostInfoSeq, bGetAll
iCrtPostInfoSeq = 0
bGetAll = False

Dim aOptName_333, aOptName_24567, aOptName_other
aOptName_333 = Array(Array(OPT_SEQ_333, "333", "3阶", "3速"), _
                        Array(OPT_SEQ_3bf, "3盲", "333bf", "3bf", "3阶盲拧", "3bld") , _
                        Array(OPT_SEQ_3oh, "3单", "333oh", "3oh", "3阶单手", "3one") , _
                        Array(OPT_SEQ_3fm, "333fm", "3fm"))

aOptName_24567 = Array(Array(OPT_SEQ_222, "222","2阶","2速"), _
                        Array(OPT_SEQ_444, "444","4阶","4速"), _
                        Array(OPT_SEQ_555, "555","5阶","5速"), _
                        Array(OPT_SEQ_666, "666","6阶","6速"), _
                        Array(OPT_SEQ_777, "777","7阶","7速"), _
                        Array(OPT_SEQ_mega, "5魔", "5魔方"))

aOptName_other = Array(Array(OPT_SEQ_3oh, "oh"), _
                        Array(OPT_SEQ_3fm, "最少步", "最小步", "fm"), _
                        Array(OPT_SEQ_mega, "mega", "megaminx"), _
                        Array(OPT_SEQ_py, "塔", "金字塔", "py", "pyramid", "pyram"), _
                        Array(OPT_SEQ_sq, "sq", "sq1", "sq-1"), _
                        Array(OPT_SEQ_clk, "表", "魔表", "clock", "clk"), _
                        Array(OPT_SEQ_sk, "sk", "斜转", "skewb"))

Sub getAllResultInfo()
    Call setInnerHtml(ID_GET_FUNCTION, "正在获取...")
    Call vaAllValidResultInfo.ResetArray()
    Call vaAllInvalidResultInfo.ResetArray()
    Call vaAllValidResultInfo.SetPreBound(vaAllPostInfo.Bound)
    Call vaAllInvalidResultInfo.SetPreBound(vaAllPostInfo.Bound)

    iMaxPostInfoSeq = vaAllPostInfo.Bound
    bGetAll = True
    Call getNextResultInfo()
End Sub

        Sub getNextResultInfo()
            Set oCrtPostInfo = vaAllPostInfo.V(iCrtPostInfoSeq)
            Call setInnerHtml(ID_CRT_POST_NUM, oCrtPostInfo.PostNum)
            Call vaAllOptLocationInfo.ResetArray()
            iCrtPostInfoSeq = iCrtPostInfoSeq + 1
            idTimer = window.setTimeout("getResultInfo()", 0, "VBScript")
        End Sub

        Sub getResultInfo()
            window.clearTimeout(idTimer)

            Dim aPostMsgWord
            aPostMsgWord = getPostMsgWordsArray()
            Call checkPostMsgWords(aPostMsgWord)
        End Sub

        Function getPostMsgWordsArray()
            window.clearTimeout(idTimer)

            Dim aPostMsgLine, iMaxLine
            aPostMsgLine = Split(oCrtPostInfo.PostMsg, VbCrlf)
            iMaxLine = UBound(aPostMsgLine)

            Dim i, j, aPostMsgWord()
            For i = 0 To iMaxLine
                ReDim Preserve aPostMsgWord(i)
                aPostMsgWord(i) = Split(aPostMsgLine(i), " ")
            Next

            getPostMsgWordsArray = aPostMsgWord
        End Function

        Sub checkPostMsgWords(aPostMsgWord)
            Dim iMaxLine
            iMaxLine = UBound(aPostMsgWord)

            Dim oEnd
            Set oEnd = New OptLocationInfo
            oEnd.OptSeq = -1
            oEnd.OptLocation = Array(iMaxLine, UBound(aPostMsgWord(iMaxLine)) + 1)
            vaAllOptLocationInfo.Append(oEnd)

            Dim i, j, sWord, iNextOptSeq
            For i = UBound(aPostMsgWord) To 0 Step -1
                For j = UBound(aPostMsgWord(i)) To 0 Step -1
                    sWord = aPostMsgWord(i)(j)
                    iNextOptSeq = checkWordAndGetOptSeq(sWord)
                    If iNextOptSeq <> "" Then
                        Dim oTmp
                        Set oTmp = New OptLocationInfo
                        oTmp.OptSeq = iNextOptSeq
                        oTmp.OptLocation = Array(i, j)
                        vaAllOptLocationInfo.Append(oTmp)
                    End If
                Next
            Next

            If vaAllOptLocationInfo.Bound > 0 Then
                Dim k
                For k = vaAllOptLocationInfo.Bound To 1 Step -1
                    Call handleResultText(aPostMsgWord, vaAllOptLocationInfo.V(k), vaAllOptLocationInfo.V(k - 1))
                Next
            Else
                Call addInvalidResultInfo("", oCrtPostInfo.PostMsg, "")
            End If

            If iCrtPostInfoSeq <= iMaxPostInfoSeq Then
                Call getNextResultInfo()
            Else
                Call completeGetAllResultInfo()
            End If
        End Sub

                Function checkWordAndGetOptSeq(sWord)
                    Dim iOptSeq
                    If Len(sWord) <= 8 _
                            And InStr(sWord, ".") = 0 _
                            And InStr(sWord, "'") = 0 Then
                        Call replaceCharForOptName(sWord)
                        iOptSeq = checkOptNameGroup(sWord)
                    End If
                    checkWordAndGetOptSeq = iOptSeq
                End Function

                Function checkOptNameGroup(str)
                    Dim firstChar, optName
                    firstChar = Mid(str, 1, 1)
                    optName = ""
                    If isNumeric(firstChar) Then
                        If firstChar = "3" Then
                            optName = checkOptNameArray(str, aOptName_333)
                        Else
                            optName = checkOptNameArray(str, aOptName_24567)
                        End If
                    Else
                        optName = checkOptNameArray(str, aOptName_other)
                    End If
                    checkOptNameGroup = optName
                End Function

                Function checkOptNameArray(str, optNameGroup)
                    Dim i, optName
                    For i = 0 To UBound(optNameGroup)
                        optName = checkOptNameStr(str, optNameGroup(i))
                        If optName <> "" Then checkOptNameArray = optName : Exit Function
                    Next
                    checkOptNameArray = ""
                End Function

                Function checkOptNameStr(str, optNameArray)
                    Dim i
                    For i = 1 To UBound(optNameArray)
                        If str = optNameArray(i) Then
                            checkOptNameStr = optNameArray(0)
                            Exit Function
                        End If
                    Next
                    checkOptNameStr = ""
                End Function

                Sub handleResultText(aPostMsgWord, oPrevOptLocationInfo, oNextOptLocationInfo)
                    Dim sResultText
                    sResultText = getResultText(aPostMsgWord, oPrevOptLocationInfo.OptLocation, oNextOptLocationInfo.OptLocation)

                    Call addValidResultInfo(oPrevOptLocationInfo.OptSeq, sResultText)
                End Sub

                Function getResultText(aPostMsgWord, aPrevOptLocation, aNextOptLocation)
                    Dim iLineHead, iWordHead, iLineEnd, iWordEnd, sTmp, aWord, iHead, iEnd
                    iLineHead = aPrevOptLocation(0)
                    iWordHead = aPrevOptLocation(1)
                    iLineEnd = aNextOptLocation(0)
                    iWordEnd = aNextOptLocation(1)
                    sTmp = ""

                    If iWordHead < UBound(aPostMsgWord(iLineHead)) Then
                        iWordHead = iWordHead + 1
                    Else
                        iLineHead = iLineHead + 1
                        iWordHead = 0
                    End If

                    If iWordEnd > 0 Then
                        iWordEnd = iWordEnd - 1
                    Else
                        iLineEnd = iLineEnd - 1
                        iWordEnd = UBound(aPostMsgWord(iLineEnd))
                    End If

                    Dim i, j
                    For i = iLineHead To iLineEnd
                        aWord = aPostMsgWord(i)
                        If i = iLineHead Then
                            iHead = iWordHead
                            isAnotherLine = False
                        Else
                            iHead = 0
                            isAnotherLine = True
                        End If

                        If i = iLineEnd Then
                            iEnd = iWordEnd
                        Else
                            iEnd = UBound(aWord)
                        End If

                        For j = iHead To iEnd
                            If isAnotherLine And (j = 0) Then
                                sTmp = sTmp & VbCrlf & aWord(j)
                            Else
                                sTmp = sTmp & " " & aWord(j)
                            End If
                        Next
                    Next
                    getResultText = LTrim(sTmp)
                End Function

        Sub addValidResultInfo(iOptSeq, sResultText)
            If sResultText = "" Then Call addInvalidResultInfo(iOptSeq, "", "") : Exit Sub

            '//check the same owner and opt
            'If checkOwnerAndOptIsExist(postUser, iOptSeq) Then Exit Sub

            '//get pure results
            Dim aPureResultInfo, sPureResults, isValid
            aPureResultInfo = getPureResult(Replace(sResultText, VbCrlf, " "), iOptSeq)
            sPureResults = aPureResultInfo(0)
            isValid = aPureResultInfo(1)

            If Not isValid Then Call addInvalidResultInfo(iOptSeq, sResultText, sPureResults) : Exit Sub


            Dim oNew : Set oNew = New ValidResultInfo
            oNew.PostNum = oCrtPostInfo.PostNum
            oNew.ResultOwner = oCrtPostInfo.PostUser
            oNew.ResultOptSeq = iOptSeq
            oNew.ResultText = sResultText
            oNew.PureResults = sPureResults

            Call getSortedResult(sPureResults, iOptSeq, oNew)

            Call vaAllValidResultInfo.Append(oNew)
            Call vaOptInfo.V(iOptSeq).CountPlus()
        End Sub

        Sub addInvalidResultInfo(iOptSeq, sResultText, sPureResults)
            Dim oNew : Set oNew = New InvalidResultInfo
            oNew.PostNum = oCrtPostInfo.PostNum
            oNew.ResultOwner = oCrtPostInfo.PostUser
            oNew.Resultoptseq = iOptSeq
            oNew.ResultText = sResultText
            oNew.PureResults = sPureResults
            Call vaAllInvalidResultInfo.Append(oNew)
        End Sub

        Function checkOwnerAndOptIsExist(sOwner, iOptSeq)
            If vaAllValidResultInfo.Bound = -1 Then
                checkOwnerAndOptIsExist = False
                Exit Function
            End If

            Dim i, oResult
            For i = 0 To vaAllValidResultInfo.Bound
                Set oResult = vaAllValidResultInfo.V(i)
                If oResult.ResultOwner = sOwner And oResult.ResultOptSeq = iOptSeq Then
                    checkOwnerAndOptIsExist = True
                    Exit Function
                End If
            Next

            checkOwnerAndOptIsExist = False
        End Function

        Sub enableElementAfterGetResultInfo()
            Call enableElement(ID_SELECT_VALID_RESULTS)
            Call enableElement(ID_SELECT_INVALID_RESULTS)
            Call enableElement(ID_SELECT_BR_RESULTS)
            Call enableElement(ID_BUTTON_ADD_NEW_RESULT_INFO)
            Call enableElement(ID_BUTTON_REMOVE_VALID_RESULT)
            Call enableElement(ID_TEXTAREA_SHOW_RESULT_POST_NUM)
            Call enableElement(ID_TEXTAREA_SHOW_RESULT_POST_USER)
            Call enableElement(ID_TEXTAREA_SHOW_RESULT_OPT_NAME)
            Call enableElement(ID_TEXTAREA_SHOW_RESULT_BEST_RESULT)
            Call enableElement(ID_TEXTAREA_SHOW_RESULT_IS_BEST_BR)
            Call enableElement(ID_TEXTAREA_SHOW_RESULT_AVG_RESULT)
            Call enableElement(ID_TEXTAREA_SHOW_RESULT_IS_AVG_BR)
            Call enableElement(ID_TEXTAREA_SHOW_RESULT_RESULT_TEXT)
            Call enableElement(ID_TEXTAREA_SHOW_RESULT_PURE_RESULTS)
            Call enableElement(ID_BUTTON_SUBMIT_NEW_RESULT_INFO)
            Call enableElement(ID_BUTTON_SELECT_NEXT_INVALID_RESULT)
            Call enableElement(ID_BUTTON_SELECT_PREV_INVALID_RESULT)
            Call enableElement(ID_BUTTON_SORT_ALL_RESULTS)
        End Sub

        Sub completeGetAllResultInfo()
            Call showAllResultInfo()
            Call enableElementAfterGetResultInfo()
            Call setInnerHtml(ID_GET_FUNCTION, "已完成")
            bGetAll = False
        End Sub
        