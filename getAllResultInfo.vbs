'*************************************************
'****split str and get Opt info.
'*************************************************
Dim vaAllValidResultInfo : Set vaAllValidResultInfo = New VariableArray
Dim vaAllInvalidResultInfo : Set vaAllInvalidResultInfo = New VariableArray
Dim uPostInfoTxtPath : uPostInfoTxtPath = CrtPath & "\tmpFiles\ResultInfo.txt"
Dim uNoPostInfoTxtPath : uNoPostInfoTxtPath = CrtPath & "\tmpFiles\NoResultInfo.txt"
Dim aAllRecords(13)

Const ID_CRT_POSTNUM = "crt_postNum"

'Dim OptName_333 : Set OptName_333 = New OptName ： OptName_333.OfficialName = "三阶" : OptName_333.OtherNames = Array()
'Dim OptName_444 : Set OptName_444 = New OptName ： OptName_444.OfficialName = "四阶" : OptName_444.OtherNames = Array()
'Dim OptName_555 : Set OptName_555 = New OptName ： OptName_555.OfficialName = "五阶" : OptName_555.OtherNames = Array()
'Dim OptName_222 : Set OptName_222 = New OptName ： OptName_222.OfficialName = "二阶" : OptName_222.OtherNames = Array()
'Dim OptName_3bf : Set OptName_3bf = New OptName ： OptName_3bf.OfficialName = "三盲" : OptName_3bf.OtherNames = Array()
'Dim OptName_3oh : Set OptName_3oh = New OptName ： OptName_3oh.OfficialName = "三单" : OptName_3oh.OtherNames = Array()
'Dim OptName_3fm : Set OptName_3fm = New OptName ： OptName_3fm.OfficialName = "最少步" : OptName_3fm.OtherNames = Array()
'Dim OptName_mega : Set OptName_mega = New OptName ： OptName_mega.OfficialName = "五魔" : OptName_mega.OtherNames = Array()
'Dim OptName_py : Set OptName_py = New OptName ： OptName_py.OfficialName = "金字塔" : OptName_py.OtherNames = Array()
'Dim OptName_sq : Set OptName_sq = New OptName ： OptName_sq.OfficialName = "SQ-1" : OptName_sq.OtherNames = Array()
'Dim OptName_clk : Set OptName_clk = New OptName ： OptName_clk.OfficialName = "魔表" : OptName_clk.OtherNames = Array()
'Dim OptName_sk : Set OptName_sk = New OptName ： OptName_sk.OfficialName = "斜转" : OptName_sk.OtherNames = Array()
'Dim OptName_666 : Set OptName_666 = New OptName ： OptName_666.OfficialName = "六阶" : OptName_666.OtherNames = Array()
'Dim OptName_777 : Set OptName_777 = New OptName ： OptName_777.OfficialName = "七阶" : OptName_777.OtherNames = Array()

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

Function getAllResultInfo()
    Call loadAllRecords()
    Call loadAllOptInfo()
    Call vaAllValidResultInfo.ResetArray()
    Call vaAllInvalidResultInfo.ResetArray()
    Call vaAllValidResultInfo.SetPreBound(vaAllPostInfo.Bound)
    Call vaAllInvalidResultInfo.SetPreBound(vaAllPostInfo.Bound)
    Dim i, optName
    For i = 0 To vaAllPostInfo.Bound
        Call setInnerHtml(ID_CRT_POSTNUM, vaAllPostInfo.V(i).PostNum)
        Call getOptName(vaAllPostInfo.V(i))
    Next
    Call saveValidResultInfoToTxt()
    Call saveInValidResultInfoToTxt()
    'MsgBox("Get all result info done!")
    Call showAllResultInfo()

    MsgBox("getAllResultInfo done!")
End Function

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

        Sub saveValidResultInfoToTxt()
            initTxtFile(uPostInfoTxtPath)
            Dim oTxt, i, obj
            Set oTxt = Fso.OpenTextFile(uPostInfoTxtPath, 8, False, True)

            For i = 0 To vaAllValidResultInfo.Bound
                Set obj = vaAllValidResultInfo.V(i)
                oTxt.WriteLine(obj.PostNum)
                oTxt.WriteLine(obj.ResultOwner)
                oTxt.WriteLine(getOptFullNameBySeq(obj.ResultOptSeq))
                oTxt.WriteLine(obj.ResultText)
                oTxt.WriteLine(obj.PureResults)
                oTxt.WriteLine(obj.SortedResults)
                oTxt.WriteLine(obj.BestResult)
                oTxt.WriteLine(obj.AvgResult)
                oTxt.WriteLine(obj.IsBestBr)
                oTxt.WriteLine(obj.IsAvgBr)
                oTxt.WriteLine()
            Next

            oTxt.Close
            Set oTxt = Nothing
        End Sub

        Sub saveInValidResultInfoToTxt()
            initTxtFile(uNoPostInfoTxtPath)
            Dim oTxt, i, obj
            Set oTxt = Fso.OpenTextFile(uNoPostInfoTxtPath, 8, False, True)

            For i = 0 To vaAllInvalidResultInfo.Bound
                Set obj = vaAllInvalidResultInfo.V(i)
                oTxt.WriteLine(obj.PostNum)
                oTxt.WriteLine(obj.ResultOwner)
                oTxt.WriteLine(getOptFullNameBySeq(obj.ResultOptSeq))
                oTxt.WriteLine(obj.ResultText)
                oTxt.WriteLine(obj.PureResults)
                oTxt.WriteLine()
            Next

            oTxt.Close
            Set oTxt = Nothing
        End Sub

        Sub getOptName(object)
            Dim aTmp, i, optSeqLast, optSeqNext, seqLast, seqNext, optCount
            optSeqLast = ""
            seqLast = 0
            optCount = 0
            aTmp = Split(object.PostMsg, " ")
            For i = 0 To UBound(aTmp)
                If Len(aTmp(i)) <= 8 And InStr(aTmp(i), ".") = 0 Then
                    Call replaceCharForOptName(aTmp(i))
                    optSeqNext = checkOptNameGroup(aTmp(i))
                    seqNext = i

                    If optSeqNext <> "" Then
                        optCount = optCount + 1
                        If optSeqLast <> "" Then
                            Call addValidResultInfo(object.PostNum, object.PostUser, optSeqLast, aTmp, seqLast, seqNext - 1, optCount)
                        End If
                        optSeqLast = optSeqNext
                        seqLast = seqNext
                    End If

                End If
            Next

            If optSeqLast <> "" Then
                Call addValidResultInfo(object.PostNum, object.PostUser, optSeqLast, aTmp, seqLast, UBound(aTmp), optCount)
            Else
                Call addInvalidResultInfo(object.PostNum, object.PostUser, "", object.PostMsg, "")
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

        Sub addValidResultInfo(postNum, postUser, optSeq, aPostMsg, seqLast, seqNext, optCount)
            Dim resultText

            If seqLast < seqNext Then
                resultText = getResultText(aPostMsg, seqLast + 1, seqNext)
            '//opt name is at last position
            ElseIf (seqLast = UBound(aPostMsg)) And (optCount = 1) Then
                resultText = getResultText(aPostMsg, 0, seqLast - 1)
            End If

            If resultText = "" Then Call addInvalidResultInfo(postNum, postUser, optSeq, "", "") : Exit Sub

            '//check the same owner and opt
            'If checkOwnerAndOptIsExist(postUser, optSeq) Then Exit Sub

            '//get pure results
            Dim aPureResultInfo, pureResults, isValid
            aPureResultInfo = getPureResult(resultText, optSeq)
            pureResults = aPureResultInfo(0)
            isValid = aPureResultInfo(1)

            If Not isValid Then Call addInvalidResultInfo(postNum, postUser, optSeq, resultText, pureResults) : Exit Sub


            Dim oNew : Set oNew = New ValidResultInfo
            oNew.PostNum = postNum
            oNew.ResultOwner = postUser
            oNew.ResultOptSeq = optSeq
            oNew.ResultText = resultText
            oNew.PureResults = pureResults

            Call getSortedResult(pureResults, optSeq, oNew)

            Call vaAllValidResultInfo.Append(oNew)
            Call vaOptInfo.V(optSeq).CountPlus()
        End Sub

        Function getResultText(aPostMsg, minSeq, maxSeq)
            Dim i, sTmp
            sTmp = ""
            For i = minSeq To maxSeq
                sTmp = sTmp & " " & aPostMsg(i)
            Next
            getResultText = LTrim(sTmp)
        End Function

        Sub addInvalidResultInfo(postNum, postUser, optSeq, resultText, pureResults)
            Dim oNew : Set oNew = New InvalidResultInfo
            oNew.PostNum = postNum
            oNew.ResultOwner = postUser
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
        