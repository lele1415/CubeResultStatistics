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

Dim vaIgnoreUsers : Set vaIgnoreUsers = New VariableArray
Dim aAllRecords(13)
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
