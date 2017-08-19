'*************************************************
'****split str and get Opt info.
'*************************************************
Dim vaAllValidResultInfo : Set vaAllValidResultInfo = New VariableArray
Dim vaAllInvalidResultInfo : Set vaAllInvalidResultInfo = New VariableArray
Dim uPostInfoTxtPath : uPostInfoTxtPath = CrtPath & "\tmpFiles\ResultInfo.txt"
Dim uNoPostInfoTxtPath : uNoPostInfoTxtPath = CrtPath & "\tmpFiles\NoResultInfo.txt"

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

Const OPT_333 = "三阶"
Const OPT_444 = "四阶"
Const OPT_555 = "五阶"
Const OPT_222 = "二阶"
Const OPT_3bf = "三盲"
Const OPT_3oh = "三单"
Const OPT_3fm = "最少步"
Const OPT_mega = "五魔"
Const OPT_py = "金字塔"
Const OPT_sq = "SQ-1"
Const OPT_clk = "魔表"
Const OPT_sk = "斜转"
Const OPT_666 = "六阶"
Const OPT_777 = "七阶"

Dim aOptName_333, aOptName_24567, aOptName_other
aOptName_333 = Array(Array(OPT_333, "333", "3阶", "3速"), _
                        Array(OPT_3bf, "3盲", "333bf", "3bf", "3阶盲拧", "3bld") , _
                        Array(OPT_3oh, "3单", "333oh", "3oh", "3阶单手", "3one") , _
                        Array(OPT_3fm, "333fm", "3fm"))

aOptName_24567 = Array(Array(OPT_222, "222","2阶","2速"), _
                        Array(OPT_444, "444","4阶","4速"), _
                        Array(OPT_555, "555","5阶","5速"), _
                        Array(OPT_666, "666","6阶","6速"), _
                        Array(OPT_777, "777","7阶","7速"), _
                        Array(OPT_mega, "5魔", "5魔方"))

aOptName_other = Array(Array(OPT_3oh, "oh"), _
                        Array(OPT_3fm, "最少步", "最小步", "fm"), _
                        Array(OPT_mega, "mega", "megaminx"), _
                        Array(OPT_py, "塔", "金字塔", "py", "pyramid", "pyram"), _
                        Array(OPT_sq, "sq", "sq1", "sq-1"), _
                        Array(OPT_clk, "表", "魔表", "clock", "clk"), _
                        Array(OPT_sk, "sk", "斜转", "skewb"))

Function getAllResultInfo()
    Call vaAllValidResultInfo.ResetArray()
    Call vaAllInvalidResultInfo.ResetArray()
    Dim i, optName
    For i = 0 To vaAllPostInfo.Bound
        Call getOptName(vaAllPostInfo.V(i))
    Next
    Call saveResultInfoToTxt(uPostInfoTxtPath, vaAllValidResultInfo)
    Call saveResultInfoToTxt(uNoPostInfoTxtPath, vaAllInvalidResultInfo)

    MsgBox("getAllResultInfo done!")
End Function

        Sub saveResultInfoToTxt(uTxtPath, vaObject)
            initTxtFile(uTxtPath)
            Dim oTxt, i
            Set oTxt = Fso.OpenTextFile(uTxtPath, 8, False, True)

            For i = 0 To vaObject.Bound
                oTxt.WriteLine(vaObject.V(i).PostNum)
                oTxt.WriteLine(vaObject.V(i).ResultOwner)
                oTxt.WriteLine(vaObject.V(i).ResultOpt)
                oTxt.WriteLine(vaObject.V(i).ResultText)
                oTxt.WriteLine(vaObject.V(i).PureResults)
                oTxt.WriteLine()
            Next

            oTxt.Close
            Set oTxt = Nothing
        End Sub

        Sub getOptName(object)
            Dim aTmp, i, j, optNameLast, optNameNext, seqLast, seqNext, optCount
            optNameLast = ""
            seqLast = 0
            optCount = 0
            aTmp = Split(object.PostMsg, " ")
            For i = 0 To UBound(aTmp)
                If Len(aTmp(i)) <= 8 And InStr(aTmp(i), ".") = 0 Then
                    Call replaceCharForOptName(aTmp(i))
                    optNameNext = checkOptNameGroup(aTmp(i))
                    seqNext = i

                    If optNameNext <> "" Then
                        optCount = optCount + 1
                        If optNameLast <> "" Then
                            Call addValidResultInfo(object.PostNum, object.PostUser, optNameLast, aTmp, seqLast, seqNext - 1, optCount)
                        End If
                        optNameLast = optNameNext
                        seqLast = seqNext
                    End If

                End If
            Next

            If optNameLast <> "" Then
                Call addValidResultInfo(object.PostNum, object.PostUser, optNameLast, aTmp, seqLast, UBound(aTmp), optCount)
            Else
                Call addInvalidResultInfo(object.PostNum, object.PostUser, optNameLast, object.PostMsg, "")
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

        Sub addValidResultInfo(postNum, postUser, optName, aPostMsg, seqLast, seqNext, optCount)
            Dim i, resultText

            If seqLast < seqNext Then
                resultText = getResultText(aPostMsg, seqLast + 1, seqNext)
            '//opt name is at last position
            ElseIf (seqLast = UBound(aPostMsg)) And (optCount = 1) Then
                resultText = getResultText(aPostMsg, 0, seqLast - 1)
            End If

            If resultText = "" Then Call addInvalidResultInfo(postNum, postUser, optName, "", "") : Exit Sub

            Dim aTmp
            aTmp = getPureResult(resultText, optName)

            If Not aTmp(1) Then Call addInvalidResultInfo(postNum, postUser, optName, resultText, aTmp(0)) : Exit Sub

            Dim oNew : Set oNew = New ValidResultInfo
            oNew.PostNum = postNum
            oNew.ResultOwner = postUser
            oNew.ResultOpt = optName
            oNew.ResultText = resultText
            oNew.PureResults = aTmp(0)
            Call vaAllValidResultInfo.Append(oNew)
        End Sub

        Function getResultText(aPostMsg, minSeq, maxSeq)
            Dim i, sTmp
            sTmp = ""
            For i = minSeq To maxSeq
                sTmp = sTmp & " " & aPostMsg(i)
            Next
            getResultText = LTrim(sTmp)
        End Function

        Sub addInvalidResultInfo(postNum, postUser, optName, resultText, pureResults)
            Dim oNew : Set oNew = New InvalidResultInfo
            oNew.PostNum = postNum
            oNew.ResultOwner = postUser
            oNew.ResultOpt = optName
            oNew.ResultText = resultText
            oNew.PureResults = pureResults
            Call vaAllInvalidResultInfo.Append(oNew)
        End Sub
        