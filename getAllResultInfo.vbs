'*************************************************
'****split str and get Opt info.
'*************************************************
Dim vaAllValidResultInfo : Set vaAllValidResultInfo = New VariableArray
Dim vaAllInvalidResultInfo : Set vaAllInvalidResultInfo = New VariableArray
Dim aAllRecords(13)

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
Const ID_BUTTON_SELECT_LAST_INVALID_RESULT = "select_last_invalid_result"
Const ID_BUTTON_SORT_ALL_RESULTS = "sort_all_results"

Const OPT_FULL_NAME_333 = "三阶"
Const OPT_FULL_NAME_444 = "四阶"
Const OPT_FULL_NAME_555 = "五阶"
Const OPT_FULL_NAME_222 = "二阶"
Const OPT_FULL_NAME_3bf = "三盲"
Const OPT_FULL_NAME_3oh = "三单"
Const OPT_FULL_NAME_3fm = "最少步"
Const OPT_FULL_NAME_mega = "五魔"
Const OPT_FULL_NAME_py = "金字塔"
Const OPT_FULL_NAME_sq = "SQ-1"
Const OPT_FULL_NAME_clk = "魔表"
Const OPT_FULL_NAME_sk = "斜转"
Const OPT_FULL_NAME_666 = "六阶"
Const OPT_FULL_NAME_777 = "七阶"

Const OPT_SEQ_333 = 0
Const OPT_SEQ_444 = 1
Const OPT_SEQ_555 = 2
Const OPT_SEQ_222 = 3
Const OPT_SEQ_3bf = 4
Const OPT_SEQ_3oh = 5
Const OPT_SEQ_3fm = 6
Const OPT_SEQ_mega = 7
Const OPT_SEQ_py = 8
Const OPT_SEQ_sq = 9
Const OPT_SEQ_clk = 10
Const OPT_SEQ_sk = 11
Const OPT_SEQ_666 = 12
Const OPT_SEQ_777 = 13

Dim vaOptInfo : Set vaOptInfo = New VariableArray
Dim oCrtPostInfo, iCrtPostInfoSeq, iMaxPostInfoSeq
iCrtPostInfoSeq = 0

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
    Call loadAllRecords()
    Call loadAllOptInfo()
    Call vaAllValidResultInfo.ResetArray()
    Call vaAllInvalidResultInfo.ResetArray()
    Call vaAllValidResultInfo.SetPreBound(vaAllPostInfo.Bound)
    Call vaAllInvalidResultInfo.SetPreBound(vaAllPostInfo.Bound)

    iMaxPostInfoSeq = vaAllPostInfo.Bound
    Call getNextResultInfo()
End Sub

        Sub getNextResultInfo()
            Set oCrtPostInfo = vaAllPostInfo.V(iCrtPostInfoSeq)
            Call setInnerHtml(ID_CRT_POST_NUM, oCrtPostInfo.PostNum)
            iCrtPostInfoSeq = iCrtPostInfoSeq + 1
            idTimer = window.setTimeout("getOptName()", 0, "VBScript")
        End Sub

        Sub loadAllOptInfo()
            Call vaOptInfo.SetPreBound(14)
            Call addOptInfo(OPT_FULL_NAME_333, OPT_SEQ_333, 5)
            Call addOptInfo(OPT_FULL_NAME_444, OPT_SEQ_444, 5)
            Call addOptInfo(OPT_FULL_NAME_555, OPT_SEQ_555, 5)
            Call addOptInfo(OPT_FULL_NAME_222, OPT_SEQ_222, 5)
            Call addOptInfo(OPT_FULL_NAME_3bf, OPT_SEQ_3bf, 3)
            Call addOptInfo(OPT_FULL_NAME_3oh, OPT_SEQ_3oh, 5)
            Call addOptInfo(OPT_FULL_NAME_3fm, OPT_SEQ_3fm, 3)
            Call addOptInfo(OPT_FULL_NAME_mega, OPT_SEQ_mega, 5)
            Call addOptInfo(OPT_FULL_NAME_py, OPT_SEQ_py, 5)
            Call addOptInfo(OPT_FULL_NAME_sq, OPT_SEQ_sq, 5)
            Call addOptInfo(OPT_FULL_NAME_clk, OPT_SEQ_clk, 5)
            Call addOptInfo(OPT_FULL_NAME_sk, OPT_SEQ_sk, 5)
            Call addOptInfo(OPT_FULL_NAME_666, OPT_SEQ_666, 3)
            Call addOptInfo(OPT_FULL_NAME_777, OPT_SEQ_777, 3)
        End Sub

        Sub addOptInfo(optFullName, optSeq, optNeedNum)
            Dim aRecordInfo, iBestRecord, iAvgRecord
            aRecordInfo = Split(aAllRecords(optSeq))
            iBestRecord = checkColonAndFormat(aRecordInfo(1))
            iAvgRecord = checkColonAndFormat(aRecordInfo(2))

            Dim optInfo : Set optInfo = New OptionInfo
            optInfo.FullName = optFullName
            optInfo.Seq = optSeq
            optInfo.NeedNum = optNeedNum
            optInfo.BestRecord = iBestRecord
            optInfo.AvgRecord = iAvgRecord
            vaOptInfo.Append(optInfo)
        End Sub

                Function checkColonAndFormat(iResult)
                    If InStr(iResult, ":") Then
                        checkColonAndFormat = formatResultStr(iResult, "")
                    Else
                        checkColonAndFormat = iResult
                    End If
                End Function

        Sub loadAllRecords()
            Dim oTxt, count
            Set oTxt = Fso.OpenTextFile(uRecordTxt, 1, False, True)
            count = 0

            Dim sLine
            Do Until oTxt.AtEndOfStream
                sLine = oTxt.ReadLine
                If Trim(sLine) <> "" Then
                    aAllRecords(count) = sLine
                    count = count + 1
                End If
            Loop

            oTxt.Close
            Set oTxt = Nothing
        End Sub

        Function getOptFullNameBySeq(iSeq)
            If iSeq <> "" Then
                getOptFullNameBySeq = vaOptInfo.V(iSeq).FullName
            Else
                getOptFullNameBySeq = ""
            End If
        End Function

        Sub getOptName()
            window.clearTimeout(idTimer)

            Dim aPostMsgLine, iMaxLine, iOptSeqLast, iOptSeqNext, aOptLastLocation, aOptNextLocation
            iOptSeqLast = ""
            aOptLastLocation = Array(0, 0)
            aOptNextLocation = Array(0, 0)
            aPostMsgLine = Split(oCrtPostInfo.PostMsg, VbCrlf)
            iMaxLine = UBound(aPostMsgLine)

            Dim i, j, aPostMsgWord()
            For i = 0 To iMaxLine
                ReDim Preserve aPostMsgWord(i)
                aPostMsgWord(i) = Split(aPostMsgLine(i), " ")
                For j = 0 To UBound(aPostMsgWord(i))
                    Dim sWord
                    sWord = aPostMsgWord(i)(j)
                    If Len(sWord) <= 8 _
                            And InStr(sWord, ".") = 0 _
                            And InStr(sWord, "'") = 0 Then
                        Call replaceCharForOptName(sWord)
                        iOptSeqNext = checkOptNameGroup(sWord)

                        If iOptSeqNext <> "" Then
                            aOptNextLocation = Array(i, j)
                            If iOptSeqLast <> "" Then
                                Call addValidResultInfo(iOptSeqLast, iMaxLine, aPostMsgWord, aOptLastLocation, aOptNextLocation)
                            End If
                            iOptSeqLast = iOptSeqNext
                            aOptLastLocation = aOptNextLocation
                        End If

                    End If
                Next
            Next

            If iOptSeqLast <> "" Then
                Call addValidResultInfo(iOptSeqLast, iMaxLine, aPostMsgWord, aOptLastLocation, Array(iMaxLine, UBound(aPostMsgWord(iMaxLine)) + 1))
            Else
                Call addInvalidResultInfo("", oCrtPostInfo.PostMsg, "")
            End If

            If iCrtPostInfoSeq <= iMaxPostInfoSeq Then
                Call getNextResultInfo()
            Else
                Call showAllResultInfo()
                Call enableElementAfterGetResultInfo()
                Call setInnerHtml(ID_GET_FUNCTION, "已完成")
            End If
        End Sub

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

        Sub addValidResultInfo(optSeq, iMaxLine, aPostMsgWord, aOptLastLocation, aOptNextLocation)
            Dim resultText
            resultText = ""

            'If (aOptLastLocation(0) <> aOptNextLocation(0)) Or (aOptLastLocation(1) <> aOptNextLocation(1)) Then
                resultText = getResultText(aPostMsgWord, aOptLastLocation, aOptNextLocation)
            '//opt name is at last position
            'ElseIf (aOptLastLocation(0) = iMaxLine) And (aOptLastLocation(1) = UBound(aPostMsgWord(iMaxLine))) Then
            '    resultText = getResultText(aPostMsgWord, Array(0, -1), aOptLastLocation)
            'End If

            If resultText = "" Then Call addInvalidResultInfo(optSeq, "", "") : Exit Sub

            '//check the same owner and opt
            'If checkOwnerAndOptIsExist(postUser, optSeq) Then Exit Sub

            '//get pure results
            Dim aPureResultInfo, pureResults, isValid
            aPureResultInfo = getPureResult(Replace(resultText, VbCrlf, " "), optSeq)
            pureResults = aPureResultInfo(0)
            isValid = aPureResultInfo(1)

            If Not isValid Then Call addInvalidResultInfo(optSeq, resultText, pureResults) : Exit Sub


            Dim oNew : Set oNew = New ValidResultInfo
            oNew.PostNum = oCrtPostInfo.PostNum
            oNew.ResultOwner = oCrtPostInfo.PostUser
            oNew.ResultOptSeq = optSeq
            oNew.ResultText = resultText
            oNew.PureResults = pureResults

            Call getSortedResult(pureResults, optSeq, oNew)

            Call vaAllValidResultInfo.Append(oNew)
            Call vaOptInfo.V(optSeq).CountPlus()
        End Sub

        Function getResultText(aPostMsgWord, aOptLastLocation, aOptNextLocation)
            Dim iLineHead, iWordHead, iLineEnd, iWordEnd, sTmp, aWord, iHead, iEnd
            iLineHead = aOptLastLocation(0)
            iWordHead = aOptLastLocation(1)
            iLineEnd = aOptNextLocation(0)
            iWordEnd = aOptNextLocation(1)
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

        Sub addInvalidResultInfo(optSeq, resultText, pureResults)
            Dim oNew : Set oNew = New InvalidResultInfo
            oNew.PostNum = oCrtPostInfo.PostNum
            oNew.ResultOwner = oCrtPostInfo.PostUser
            oNew.ResultOptSeq = optSeq
            oNew.ResultText = resultText
            oNew.PureResults = pureResults
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
            Call enableElement(ID_BUTTON_SELECT_LAST_INVALID_RESULT)
            Call enableElement(ID_BUTTON_SORT_ALL_RESULTS)
        End Sub
        