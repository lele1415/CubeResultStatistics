'*************************************************
'****split str and get Opt info.
'*************************************************
Dim vaAllResultInfo : Set vaAllResultInfo = New VariableArray
Dim vaAllNoResultInfo : Set vaAllNoResultInfo = New VariableArray
Dim uPostInfoTxtPath : uPostInfoTxtPath = CrtPath & "\tmpFiles\ResultInfo.txt"
Dim uNoPostInfoTxtPath : uNoPostInfoTxtPath = CrtPath & "\tmpFiles\NoResultInfo.txt"

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
    Call vaAllResultInfo.ResetArray()
    Call vaAllNoResultInfo.ResetArray()
    Dim i, optName
    For i = 0 To vaAllPostInfo.Bound
        Call getOptName(vaAllPostInfo.V(i))
    Next
    Call saveResultInfo(uPostInfoTxtPath, vaAllResultInfo)
    Call saveResultInfo(uNoPostInfoTxtPath, vaAllNoResultInfo)

    MsgBox("getAllResultInfo done!")
End Function

        Sub saveResultInfo(uTxtPath, vaObject)
            initTxtFile(uTxtPath)
            Dim oTxt, i
            Set oTxt = Fso.OpenTextFile(uTxtPath, 8, False, True)

            For i = 0 To vaObject.Bound
                oTxt.WriteLine(vaObject.V(i).ResultOwner)
                oTxt.WriteLine(vaObject.V(i).ResultOpt)
                oTxt.WriteLine(vaObject.V(i).ResultText)
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
                            Call addResultInfo(object, optNameLast, aTmp, seqLast, seqNext - 1, optCount)
                        End If
                        optNameLast = optNameNext
                        seqLast = seqNext
                    End If

                End If
            Next

            If optNameLast <> "" Then
                Call addResultInfo(object, optNameLast, aTmp, seqLast, UBound(aTmp), optCount)
            Else
                Call addNoResultInfo(object, optNameLast)
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

        Sub addResultInfo(object, optName, aPostMsg, seqLast, seqNext, optCount)
            Dim i, resultText

            If seqLast < seqNext Then
                resultText = getResultText(aPostMsg, seqLast + 1, seqNext)
            '//opt name is at last position
            ElseIf (seqLast = UBound(aPostMsg)) And (optCount = 1) Then
                resultText = getResultText(aPostMsg, 0, seqLast - 1)
            End If

            If resultText = "" Then Call addNoResultInfo(object, optName) : Exit Sub

            Dim oNew : Set oNew = New ResultInfo
            oNew.ResultOwner = object.PostUser
            oNew.ResultOpt = optName
            oNew.ResultText = resultText
            oNew.PureResults = getPureResult(resultText, optName)
            Call vaAllResultInfo.Append(oNew)
        End Sub

        Function getResultText(aPostMsg, minSeq, maxSeq)
            Dim i, sTmp
            sTmp = ""
            For i = minSeq To maxSeq
                sTmp = sTmp & " " & aPostMsg(i)
            Next
            getResultText = LTrim(sTmp)
        End Function

        Sub addNoResultInfo(object, optName)
            Dim oNew : Set oNew = New ResultInfo
            oNew.ResultOwner = object.PostUser
            oNew.ResultOpt = optName
            oNew.ResultText = object.PostMsg
            Call vaAllNoResultInfo.Append(oNew)
        End Sub
        