Const OPT_FULL_NAME_333 = "三阶"
Const OPT_FULL_NAME_222 = "二阶"
Const OPT_FULL_NAME_444 = "四阶"
Const OPT_FULL_NAME_555 = "五阶"
Const OPT_FULL_NAME_666 = "六阶"
Const OPT_FULL_NAME_777 = "七阶"
Const OPT_FULL_NAME_3bf = "三盲"
Const OPT_FULL_NAME_3fm = "最少步"
Const OPT_FULL_NAME_3oh = "三单"
Const OPT_FULL_NAME_3wf = "脚拧"
Const OPT_FULL_NAME_mega = "五魔"
Const OPT_FULL_NAME_py = "金字塔"
Const OPT_FULL_NAME_clk = "魔表"
Const OPT_FULL_NAME_sk = "斜转"
Const OPT_FULL_NAME_sq = "SQ-1"
Const OPT_FULL_NAME_4bf = "四盲"
Const OPT_FULL_NAME_5bf = "五盲"
Const OPT_FULL_NAME_3mb = "多盲"

Const OPT_SEQ_333 = 0
Const OPT_SEQ_222 = 1
Const OPT_SEQ_444 = 2
Const OPT_SEQ_555 = 3
Const OPT_SEQ_666 = 4
Const OPT_SEQ_777 = 5
Const OPT_SEQ_3bf = 6
Const OPT_SEQ_3fm = 7
Const OPT_SEQ_3oh = 8
Const OPT_SEQ_3wf = 9
Const OPT_SEQ_mega = 10
Const OPT_SEQ_py = 11
Const OPT_SEQ_clk = 12
Const OPT_SEQ_sk = 13
Const OPT_SEQ_sq = 14
Const OPT_SEQ_4bf = 15
Const OPT_SEQ_5bf = 16
Const OPT_SEQ_3mb = 17

Dim aOptName_333, aOptName_24567, aOptName_other
aOptName_333 = Array(Array(OPT_SEQ_333, "333", "3阶", "3速"), _
                        Array(OPT_SEQ_3bf, "3盲", "333bf", "3bf", "3阶盲拧", "3bld"), _
                        Array(OPT_SEQ_3oh, "3单", "333oh", "3oh", "3阶单手", "3one") , _
                        Array(OPT_SEQ_3fm, "333fm", "3fm"), _
                        Array(OPT_SEQ_3wf, "333wf", "3wf"), _
                        Array(OPT_SEQ_3mb, "3mb"))

aOptName_24567 = Array(Array(OPT_SEQ_222, "222","2阶","2速"), _
                        Array(OPT_SEQ_444, "444","4阶","4速"), _
                        Array(OPT_SEQ_555, "555","5阶","5速"), _
                        Array(OPT_SEQ_666, "666","6阶","6速"), _
                        Array(OPT_SEQ_777, "777","7阶","7速"), _
                        Array(OPT_SEQ_mega, "5魔", "5魔方"), _
                        Array(OPT_SEQ_4bf, "4盲", "4bf", "4bld"), _
                        Array(OPT_SEQ_5bf, "5盲", "5bf", "5bld"))

aOptName_other = Array(Array(OPT_SEQ_3oh, "oh"), _
                        Array(OPT_SEQ_3fm, "最少步", "最小步", "fm"), _
                        Array(OPT_SEQ_mega, "mega", "megaminx"), _
                        Array(OPT_SEQ_py, "塔", "金字塔", "py", "pyramid", "pyram"), _
                        Array(OPT_SEQ_sq, "sq", "sq1", "sq-1"), _
                        Array(OPT_SEQ_clk, "表", "魔表", "clock", "clk"), _
                        Array(OPT_SEQ_sk, "sk", "斜转", "skewb"), _
                        Array(OPT_SEQ_3wf, "脚拧", "wf"), _
                        Array(OPT_SEQ_3mb, "多盲"))

Dim vaIgnoreUsers : Set vaIgnoreUsers = New VariableArray
Dim aAllRecords(17)
Dim vaOptInfo : Set vaOptInfo = New VariableArray

Call loadIgnoreUsers()
Call loadAllRecords()
Call loadAllOptInfo()

Sub loadIgnoreUsers()
    Dim oTxt
    Set oTxt = Fso.OpenTextFile(uIgnoreUsersTxt, 1, False, True)

    Dim sLine
    Do Until oTxt.AtEndOfStream
        sLine = Trim(oTxt.ReadLine)
        If sLine <> "" Then
            vaIgnoreUsers.Append(sLine)
        End If
    Loop

    oTxt.Close
    Set oTxt = Nothing
End Sub

Sub loadAllRecords()
    Dim oTxt, count
    Set oTxt = Fso.OpenTextFile(uRecordTxt, 1, False, True)
    count = 0

    Dim sLine
    Do Until oTxt.AtEndOfStream
        sLine = Trim(oTxt.ReadLine)
        If sLine <> "" Then
            aAllRecords(count) = sLine
            count = count + 1
        End If
    Loop

    oTxt.Close
    Set oTxt = Nothing
End Sub

Sub loadAllOptInfo()
    Call vaOptInfo.SetPreBound(17)
    Call addOptInfo(OPT_FULL_NAME_333, OPT_SEQ_333, 5)
    Call addOptInfo(OPT_FULL_NAME_222, OPT_SEQ_222, 5)
    Call addOptInfo(OPT_FULL_NAME_444, OPT_SEQ_444, 5)
    Call addOptInfo(OPT_FULL_NAME_555, OPT_SEQ_555, 5)
    Call addOptInfo(OPT_FULL_NAME_666, OPT_SEQ_666, 3)
    Call addOptInfo(OPT_FULL_NAME_777, OPT_SEQ_777, 3)
    Call addOptInfo(OPT_FULL_NAME_3bf, OPT_SEQ_3bf, 3)
    Call addOptInfo(OPT_FULL_NAME_3fm, OPT_SEQ_3fm, 3)
    Call addOptInfo(OPT_FULL_NAME_3oh, OPT_SEQ_3oh, 5)
    Call addOptInfo(OPT_FULL_NAME_3wf, OPT_SEQ_3wf, 3)
    Call addOptInfo(OPT_FULL_NAME_mega, OPT_SEQ_mega, 5)
    Call addOptInfo(OPT_FULL_NAME_py, OPT_SEQ_py, 5)
    Call addOptInfo(OPT_FULL_NAME_clk, OPT_SEQ_clk, 5)
    Call addOptInfo(OPT_FULL_NAME_sk, OPT_SEQ_sk, 5)
    Call addOptInfo(OPT_FULL_NAME_sq, OPT_SEQ_sq, 5)
    Call addOptInfo(OPT_FULL_NAME_4bf, OPT_SEQ_4bf, 3)
    Call addOptInfo(OPT_FULL_NAME_5bf, OPT_SEQ_5bf, 3)
    Call addOptInfo(OPT_FULL_NAME_3mb, OPT_SEQ_3mb, 3)
End Sub

        Sub addOptInfo(optFullName, optSeq, optNeedNum)
            Dim aRecordInfo, iBestRecord, iAvgRecord
            aRecordInfo = Split(aAllRecords(optSeq))
            iBestRecord = checkColonAndFormat(aRecordInfo(1))
            If optSeq < 15 Then
                iAvgRecord = checkColonAndFormat(aRecordInfo(2))
            End If

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
                        checkColonAndFormat = formatBrStr(iResult)
                    Else
                        checkColonAndFormat = iResult
                    End If
                End Function
