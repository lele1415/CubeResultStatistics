Dim width,height
width=CreateObject("HtmlFile").ParentWindow.Screen.AvailWidth
Window.MoveTo width-700,100
Window.ResizeTo 700,700

Dim ws, Fso
Set ws=CreateObject("wscript.shell")
Set Fso=CreateObject("Scripting.FileSystemObject")

Dim uCodeTxtPath, uOriginMsgTxtPath, uRecordTxt, uValidResultTxt
uCodeTxtPath = ws.CurrentDirectory&"\code.txt"
uOriginMsgTxtPath = ws.CurrentDirectory&"\OriginMsg.txt"
uRecordTxt = ws.CurrentDirectory&"\record.txt"
uValidResultTxt = ws.CurrentDirectory&"\validResult.txt"
uInvalidResultTxt = ws.CurrentDirectory&"\invalidResult.txt"

Dim countResultData, countNoResultData, brResultCount, iCompNum, countExcelRow, seqOfShowingResult, listOfShowingResult
countResultData = 0
countNoResultData = 0
brResultCount = 0
iCompNum = 0
countExcelRow = 0
seqOfShowingResult = 0
listOfShowingResult = ""

Dim testFlag, inputFlag, listFlag, getResultDoneFlag, checkResultNumTooMany, checkResultNumToofew
testFlag = False
inputFlag = False
listFlag = False
getResultDoneFlag = False

Dim aMessageText, aUserText, aPostNum

Dim iStartPostNum, iOriginDataArrayLength, iTmpResultCount
iStartPostNum = 0
iTmpResultCount = 0

Dim aAllPostName(), aAllUserText(), aAllOptName(), aAllMessageText(), aAllPureResults(), aAllIsValid(), aAllSortedResults()
Dim aAllRecords(13)
Dim brResultData()
ReDim brResultData(4,0)

Dim aTmpUserText_333, aTmpPureResults_333, aTmpSortedResults_333
Dim aTmpUserText_444, aTmpPureResults_444, aTmpSortedResults_444
Dim aTmpUserText_555, aTmpPureResults_555, aTmpSortedResults_555
Dim aTmpUserText_222, aTmpPureResults_222, aTmpSortedResults_222
Dim aTmpUserText_333bf, aTmpPureResults_333bf, aTmpSortedResults_333bf
Dim aTmpUserText_333oh, aTmpPureResults_333oh, aTmpSortedResults_333oh
Dim aTmpUserText_333fm, aTmpPureResults_333fm, aTmpSortedResults_333fm
Dim aTmpUserText_py, aTmpPureResults_py, aTmpSortedResults_py
Dim aTmpUserText_mega, aTmpPureResults_mega, aTmpSortedResults_mega
Dim aTmpUserText_sq, aTmpPureResults_sq, aTmpSortedResults_sq
Dim aTmpUserText_clock, aTmpPureResults_clock, aTmpSortedResults_clock
Dim aTmpUserText_sk, aTmpPureResults_sk, aTmpSortedResults_sk
Dim aTmpUserText_666, aTmpPureResults_666, aTmpSortedResults_666
Dim aTmpUserText_777, aTmpPureResults_777, aTmpSortedResults_777



Function cutStrWithKey(str, keyBegin, numBegin, keyEnd, numEnd)
    Dim tmpStr, countStart, countEnd
    tmpStr = str
    If keyBegin <> "" Then
        countStart = InStr(tmpStr, keyBegin) + numBegin
    Else
        countStart = 1
    End If
    countEnd = InStr(tmpStr, keyEnd) + numEnd
    tmpStr = Mid(tmpStr, countStart, countEnd-countStart)
    cutStrWithKey = tmpStr
End Function

Function msgboxArray(strBegin, strArray)
    Dim tmpStr
    tmpStr = strBegin
    For count = 0 To safeUBound(strArray, "msgboxArray", 1)
    	tmpStr = tmpStr & " " & strArray(count)
    Next
    'Msgbox(tmpStr)
End Function

Function getExistInArray(strArray, str)
    Dim countFinded, count
    countFinded = ""
    For count = 0 To safeUBound(strArray, "getExistInArray", 1)
        If str = strArray(count) Then
            countFinded = count
            Exit For
        End If
    Next
    getExistInArray = countFinded
End Function

Function isCharacter(character,characters)
    Dim flag
    flag = False
    If InStr(characters, character) > 0 Then
        flag = True
    End If
    isCharacter = flag
End Function

Function isEvenNum(num)
    Dim flag, tmpNum
    flag = False
    tmpNum = num
    tmpNum = tmpNum / 2
    If tmpNum = Round(tmpNum, 0) Then
        flag = True
    End If
    isEvenNum = flag
End Function

Function getAllOptName()
    tmpArray = Array("333","444","555","222","333bf","333oh","333fm","mega","py","sq","clock","sk","666","777")
    getAllOptName = tmpArray
End Function

Function getSeqInAllOptName(sOptName)
    Dim aTmp, iSeq
    aTmp = getAllOptName()
    iSeq = ""

    For i = 0 To UBound(aTmp)
        If aTmp(i) = sOptName Then
            iSeq = i
            Exit For
        End If
    Next

    If iSeq <> "" Then
        getSeqInAllOptName = iSeq
    Else
        MsgBox("Not a option name!")
    End If
End Function

Function getFinalOptName(OptName)
    Dim oldOptNameArray, newOptNameArray, finalOptName
    oldOptNameArray = getAllOptName()
    newOptNameArray = Array("三阶","四阶","五阶","二阶","三盲","三单","最少步","五魔","金字塔","SQ-1","魔表","斜转","六阶","七阶")

    For i = 0 To UBound(oldOptNameArray)
        If OptName = oldOptNameArray(i) Then
            finalOptName = newOptNameArray(i)
            Exit For
        End If
    Next

    getFinalOptName = finalOptName
End Function

Function getOptCheckStrArray(OptName)
    checkStrArray_333 = Array("333","三阶","3阶","三速","3速")
    checkStrArray_444 = Array("444","四阶","4阶","四速","4速")
    checkStrArray_555 = Array("555","五阶","5阶","五速","5速")
    checkStrArray_222 = Array("222","二阶","2阶","二速","2速")
    checkStrArray_666 = Array("666","六阶","6阶","六速","6速")
    checkStrArray_777 = Array("777","七阶","7阶","七速","7速")
    checkStrArray_333bf = Array("三盲","333bf","3bf","3bld")
    checkStrArray_333oh = Array("三单","333oh","3oh","oh","3one")
    checkStrArray_333fm = Array("最少步","最小步","333fm","3fm","fm")
    checkStrArray_mega = Array("五魔","mega","megaminx","五魔方")
    checkStrArray_py = Array("金字塔","塔","py","pyramid")
    checkStrArray_sq = Array("sq","sq1","sq-1")
    checkStrArray_clock = Array("魔表","表","clock","clk")
    checkStrArray_sk = Array("斜转","sk","skewb")

    Select Case OptName
        Case "333"
            getOptCheckStrArray = checkStrArray_333
        Case "444"
            getOptCheckStrArray = checkStrArray_444
        Case "555"
            getOptCheckStrArray = checkStrArray_555
        Case "222"
            getOptCheckStrArray = checkStrArray_222
        Case "666"
            getOptCheckStrArray = checkStrArray_666
        Case "777"
            getOptCheckStrArray = checkStrArray_777
        Case "333bf"
            getOptCheckStrArray = checkStrArray_333bf
        Case "333oh"
            getOptCheckStrArray = checkStrArray_333oh
        Case "333fm"
            getOptCheckStrArray = checkStrArray_333fm
        Case "mega"
            getOptCheckStrArray = checkStrArray_mega
        Case "py"
            getOptCheckStrArray = checkStrArray_py
        Case "sq"
            getOptCheckStrArray = checkStrArray_sq
        Case "clock"
            getOptCheckStrArray = checkStrArray_clock
        Case "sk"
            getOptCheckStrArray = checkStrArray_sk
    End Select
End Function

Function sortIntArray(aArray)
    Dim i, j, lenArray, aTmp
    aTmp = aArray
    lenArray = safeUBound(aTmp, "sortIntArray", 1)
    For i = 0 to lenArray - 1
        For j = i + 1 to lenArray
            If (aTmp(i) - aTmp(j)) > 0 Then
                swapTwoStrings aTmp(i), aTmp(j)
            End If
        Next
    Next
    sortIntArray = aTmp
End Function

Function getCharacterCountInStr(sStr, sCharacter)
    Dim iStep, sTmp, iCount
    sTmp = sStr
    iCount = 0
    Do Until Not (InStr(sTmp, sCharacter) > 0)
        sTmp = RePlace(sTmp, sCharacter, "", 1, 1)
        iCount = iCount + 1
    Loop

    getCharacterCountInStr = iCount
End Function

Function replaceCharacterInResultStr(sStr)
    Dim sTmp, sSimple, sCharacter
    sTmp = sStr
    sSimple = ""
    For searchStep = 1 To Len(sTmp)
        sCharacter = safeMid(sTmp, searchStep, 1, "replaceCharacterInResultStr")
        Select Case True
            Case IsNumeric(sCharacter)
                sSimple = sSimple & sCharacter
            Case isCharacter(sCharacter, ".。")
                sSimple = sSimple & "."
            Case isCharacter(sCharacter, ":：")
                sSimple = sSimple & ":"
            Case isCharacter(sCharacter, ",，")
                sSimple = sSimple & " "
            Case isCharacter(sCharacter, "dD")
                sSimple = sSimple & "D"
            Case isCharacter(sCharacter, "nN")
                sSimple = sSimple & "N"
            Case isCharacter(sCharacter, "fF")
                sSimple = sSimple & "F"
            Case Else
                sSimple = sSimple & " "
        End Select
    Next

    Do Until Not (InStr(sSimple, "  ") > 0)
        sSimple = RePlace(sSimple, "  ", " ")
    Loop

    replaceCharacterInResultStr = sSimple
End Function

Function formatResultStr(sStr)
    Dim sTmp, sFormat
    sTmp = sStr
    sFormat = ""

    Dim lenTmp, iPointInStr
    lenTmp = Len(sTmp)
    iPointInStr = InStr(sTmp, ".")

    If iPointInStr > 0 Then
        If lenTmp - iPointInStr > 2 Then
            sTmp = safeMid(sTmp, 1, iPointInStr + 2, "formatResultStr")
        End If
    End If

    If IsNumeric(sTmp) Then
        If InStr(sTmp, ".") > 0 And _
                Len(sTmp) >= 4 Then
            sFormat = FormatNumber(sTmp, 2, , , 0)
        End If
    ElseIf InStr(sTmp, "DNF") > 0 Or _
            InStr(sTmp, "DNS") > 0 Then
        sFormat = 9999.99
    Else
        Dim iColonInStr, iColonCount, iPointCount, sPartMin, sPartSec
        iColonInStr = InStr(sTmp, ":")
        iColonCount = getCharacterCountInStr(sTmp, ":")
        iPointCount = getCharacterCountInStr(sTmp, ".")
        'MsgBox("iColonInStr="&iColonInStr&Vblf&_
        '       "iPointInStr="&iPointInStr&Vblf&_
        '       "iColonCount="&iColonCount&Vblf&_
        '       "iPointCount="&iPointCount)

        If iColonInStr > 1 And _
                iPointInStr > 1 And _
                iColonInStr < Len(sTmp) And _
                iPointInStr < Len(sTmp) And _
                iPointInStr > iColonInStr + 1 And _
                iColonCount = 1 And _
                iPointCount = 1 Then
            sPartMin = safeMid(sTmp, 1, iColonInStr-1, "formatResultStr 111")
            sPartSec = safeMid(sTmp, iColonInStr+1, 0, "formatResultStr 222")
            'MsgBox( "formatResultStr "&Vblf&_
                    '"sPartMin="&sPartMin&Vblf&_
                    '"sPartSec="&sPartSec )
            If IsNumeric(sPartMin) And IsNumeric(sPartSec) Then
                sFormat = sPartMin * 60 + sPartSec
                sFormat = FormatNumber(sFormat, 2, , , 0)
                'MsgBox( "formatResultStr "&Vblf&_
                        '"sTmp="&sTmp )
            Else
                MsgBox( "Error: is not number"&Vblf&_
                        "sPartMin="&sPartMin&Vblf&_
                        "sPartSec="&sPartSec )
            End If
        End If
    End If

    formatResultStr = sFormat
End Function

Function formatResultStrForFm(sStr)
    Dim sTmp, sFormat
    sTmp = sStr
    sFormat = ""

    If IsNumeric(sTmp) Then
        If Len(sTmp) = 2  Then
            If IsNumeric(Mid(sTmp, 1, 1)) And _
                    IsNumeric(Mid(sTmp, 2, 1)) Then
                sFormat = sTmp
            End If
        End If
    ElseIf InStr(sTmp, "DNF") > 0 Or _
            InStr(sTmp, "DNS") > 0 Then
        sFormat = 9999.99
    End If

    formatResultStrForFm = sFormat
End Function

Function revertResult(iResult, sOptName)
    Dim sTmp
    sTmp = iResult

    If sOptName <> "333fm" Then
        Dim minNum, secNum
        If sTmp = 9999.99 Then
            sTmp = "DNF"
        ElseIf sTmp > 59.99 Then
            minNum = int(sTmp / 60)
            secNum = FormatNumber(sTmp - minNum * 60, 2, , , 0)
            If secNum < 1.00 Then
                secNum = "00" & secNum
            ElseIf secNum < 10.00 Then
                secNum = "0" & secNum
            End If
            sTmp = minNum & ":" & secNum
        ElseIf sTmp < 1.00 Then
            sTmp = "0" & sTmp
        Else
            sTmp = FormatNumber(sTmp, 2, , , 0)
        End If
        
        revertResult = sTmp
    Else
        If sTmp = 9999.99 Then sTmp = "DNF"
        revertResult = sTmp
    End If
End Function

Function replaceHtmlStr(sStr)
	Dim sTmp, iImgStart, iImgLen, iDivStart, iDivLen
	sTmp = sStr
	sTmp = RePlace(sTmp, "<br>", " ")
	sTmp = RePlace(sTmp, "&#39;", "'")
    
    Do Until Not (InStr(sTmp, "<img") > 0)
        iImgStart = InStr(sTmp, "<img")
        If iImgStart > 0 Then
            iImgLen = InStr(safeMid(sTmp, iImgStart, 0, "replaceHtmlStr 111"), ">")
            If iImgLen > 0 Then
                sTmp = RePlace(sTmp, safeMid(sTmp, iImgStart, iImgLen, "replaceHtmlStr 222"), " ")
            End If
        End If
    Loop

    Do Until Not (InStr(sTmp, "<div") > 0)
        iDivStart = InStr(sTmp, "<div")
        If iDivStart > 0 Then
            iDivLen = InStr(safeMid(sTmp, iDivStart, 0, "replaceHtmlStr 333"), "</div>") + 5
            If iDivLen > 0 Then
                sTmp = RePlace(sTmp, safeMid(sTmp, iDivStart, iDivLen, "replaceHtmlStr 444"), " ")
            End If
        End If
    Loop

    Do Until Not (InStr(sTmp, "  ") > 0)
        sTmp = RePlace(sTmp, "  ", " ")
    Loop

	replaceHtmlStr = sTmp
End Function

Function replaceStrForOptName(sStr)
    Dim sTmp
    sTmp = sStr
    sTmp = RePlace(sTmp, "[", "")
    sTmp = RePlace(sTmp, "]", "")
    sTmp = LCase(sTmp)
    replaceStrForOptName = sTmp
End Function

FUnction clearArray(strArray)
	For i = 0 To safeUBound(strArray, "clearArray", 1)
        strArray(i) = ""
    Next
End Function

Function getSeqInArrayBySearchStr(sSearchStr, aArray)
	Dim i
	For i = 0 To safeUBound(aArray, "getSeqInArrayBySearchStr", 1)
        If aArray(i) = sSearchStr Then
        	getSeqInArrayBySearchStr = i
        	Exit For
        End If
    Next
End Function

Function safeMid(sString, iStart, iLength, sFun)
	If iStart > 0 And _
		    iStart + iLength - 1 <= Len(sString) Then
		If iLength > 0 Then
            safeMid = Mid(sString, iStart, iLength)
        Else
            safeMid = Mid(sString, iStart)
        End If
    Else
        safeMid = ""
        MsgBox( "Error: safeMid ==from==>" & sFun & Vblf &_
                "sString=" & sString & Vblf &_
                "iStart=" & iStart & Vblf &_
                "iLength=" & iLength _
                )
    End If
End Function

Function safeUBound(aArray, sFun, dimension)
    If isArray(aArray) Then
        safeUBound = UBound(aArray, dimension)
    Else
        MsgBox("Error: safeUBound ==from==>" & sFun)
    End If
End Function

Function initTxtFile(FilePath)
    If Fso.FileExists(FilePath) Then
        Dim TxtFile
        Set TxtFile = Fso.getFile(FilePath)
        TxtFile.Delete
        Set TxtFile = Nothing
    End If    
    Fso.CreateTextFile FilePath, True
End Function

Function writePageCodeToTxt(responseText)
    Set oTxt = Fso.OpenTextFile(uCodeTxtPath, 8, False, True)
    oTxt.Write(responseText)
    oTxt.Close
    Set oTxt = Nothing
End Function

Function getInfoFromTxt(strSearch, strStart, numStart, strEnd)
    Dim oTxt, getStrLine, Flag
    Set oTxt = Fso.OpenTextFile(uCodeTxtPath, 1, False, True)
    Flag = False

    Do Until oTxt.AtEndOfStream
        getStrLine = oTxt.ReadLine
        Sleep(10)
        If InStr(getStrLine,strSearch) > 0 Then
            Flag = True
            Exit Do
        End If
    Loop
    oTxt.Close
    Set oTxt = Nothing
    If Flag Then
    	getStrLine = safeMid(getStrLine, InStr(getStrLine, strSearch), 0, "getInfoFromTxt(1)")
        countStart = InStr(getStrLine,strStart) + numStart
        getStrLine = safeMid(getStrLine, countStart, 0, "getInfoFromTxt(2)")
        countLength = InStr(getStrLine, strEnd) - 1
        'Msgbox(countStart & " " & countLength)
        getInfoFromTxt = safeMid(getStrLine, 1, countLength, "getInfoFromTxt(3)")
    Else
        Msgbox("Not found """ & strSearch & """")
        getInfoFromTxt = ""
    End If
End Function

Function cutArray(aOrigin, iStart, iEnd)
    'MsgBox("iStart="&iStart&" iEnd="&iEnd)
	ReDim aCuted(iEnd - iStart)
	For i = iStart To iEnd
		aCuted(i - iStart) = aOrigin(i)
	Next
	cutArray = aCuted
End Function

Function joinArrayWithSpace(aOrigin)
	Dim sJoined
	sJoined = ""
	For i = 0 To safeUBound(aOrigin, "joinArrayWithSpace", 1)
		sJoined = sJoined & " " & aOrigin(i)
	Next
	joinArrayWithSpace = Trim(sJoined)
End Function

Function coverInArray(aArray, iSeq)
    Dim lenArray
    MsgBox("111")
    lenArray = safeUBound(aArray, "coverInArray", 1)
    If iSeq < lenArray Then
        For i = iSeq To lenArray - 1
            aArray(i) = aArray(i + 1)
        Next
    End If
    MsgBox("222")
    
    ReDim Preserve aArray(lenArray - 1)
    MsgBox("333")
End Function

Function swapTwoStrings(s1, s2)
    Dim sTmp
    sTmp = s1
    s1 = s2
    s2 = sTmp
End Function



'*************************************************
'****get code from Pages, and write into text file.
'*************************************************
Function getPagesCode()
    initTxtFile(uCodeTxtPath)
    'getCode(getUrl())
    getSource(getUrl())
    setInnerHtmlById "PageNum_id", 1
    Sleep(1000)

    Dim iMaxPageNum, sUrl, i
    iMaxPageNum = getInfoFromTxt("total_page", "total_page", 12, "};")
    If IsNumeric(iMaxPageNum) Then
        For i = 2 To iMaxPageNum
            sUrl = getUrl() & "?pn=" & i
            'getCode(sUrl)
            getSource(sUrl)
            setInnerHtmlById "PageNum_id", i
            Sleep(1000)
        Next
        'Sleep(1000)
        'getMsgArrayFromCode(0)
        Msgbox("Done!")
    Else
        Msgbox("Fail! iMaxPageNum=" & iMaxPageNum)
    End If
End Function

        Sub Sleep(MSecs)  
            Dim fso
            Dim objOutputFile

            Set fso = CreateObject("Scripting.FileSystemObject") 
            If Fso.FileExists("sleeper.vbs")=False Then 
                Set objOutputFile = fso.CreateTextFile("sleeper.vbs", True) 
                objOutputFile.Write "wscript.sleep WScript.Arguments(0)" 
                objOutputFile.Close 
            End If 
             CreateObject("WScript.Shell").Run "sleeper.vbs " & MSecs,1 , True 
        End Sub

        Function getCode(url)
            Dim xmlHTTP
            Set xmlHTTP = CreateObject("MSXML2.ServerXMLHTTP")
            xmlHTTP.open "GET", url
            xmlHttp.send
            writePageCodeToTxt(xmlHttp.responseText)
            Set xmlHTTP = Nothing
        End Function



'*************************************************
'****get data array.
'*************************************************
Function getMsgArrayFromCode()
    Dim bContinue
    aPostNum = getArrayFromTxt("post_no&quot", "post_no&quot", 14, ",", -1)
    aUserText = getArrayFromTxt("alog-group=""p_author""", "_blank", 8, "</a>", -1)
    aMessageText = getArrayFromTxt("j_d_post_content", "j_d_post_content", 40, "user-hide-post-down", -32)
    iCompNum = getInfoFromTxt("<title>", "第", 1, "期")
    bContinue = True

    checkArrayForBubble()

    Dim lenMessageText, lenUserText, lenPostNum
    lenMessageText = safeUBound(aMessageText, "getMsgArrayFromCode 111", 1)
    lenUserText = safeUBound(aUserText, "getMsgArrayFromCode 222", 1)
    lenPostNum = safeUBound(aPostNum, "getMsgArrayFromCode 333", 1)
    If lenMessageText <> lenUserText _
            Or lenUserText <> lenPostNum Then
        MsgBox( "Error: getMsgArrayFromCode" & Vblf &_
                "UBound(aMessageText)=" & lenMessageText & Vblf &_
                "UBound(aUserText)=" & lenUserText & Vblf &_
                "UBound(aPostNum)=" & lenPostNum )
        bContinue = False
    Else
        iOriginDataArrayLength = lenMessageText
        initTxtFile(uOriginMsgTxtPath)
        saveOriginMsgTxt()
    End If

    If Not IsNumeric(iCompNum) Then
        iCompNum = InputBox("Input CompNum")
        Do Until IsNumeric(iCompNum)
            iCompNum = InputBox("Input CompNum")
        Loop
    End If

    If bContinue Then
        spiltStrAndGetOptInfo()
    End If
End Function

        Function getArrayFromTxt(sSearch, sCutBegin, iPlusNumToBegin, sCutEnd, iPlusNumToEnd)
            Dim oTxt, sReadLine, aGet(), iCount
            Set oTxt = Fso.OpenTextFile(uCodeTxtPath, 1, False, True)
            iCount = 0

            Do Until oTxt.AtEndOfStream
                sReadLine = oTxt.ReadLine
                If InStr(sReadLine,sSearch) > 0 Then
                    ReDim Preserve aGet(iCount)
                    iCutBegin = InStr(sReadLine, sCutBegin) + iPlusNumToBegin
                    sReadLine = safeMid(sReadLine, iCutBegin, 0, "getArrayFromTxt(1)")
                    iCutEnd = InStr(sReadLine, sCutEnd) + iPlusNumToEnd
                    aGet(iCount) = safeMid(sReadLine, 1, iCutEnd, "getArrayFromTxt(2)")
                    iCount = iCount + 1
                End If
            Loop

            oTxt.Close
            Set oTxt = Nothing
            getArrayFromTxt = aGet
        End Function

        Function checkArrayForBubble()
            For i = 0 To safeUBound(aMessageText, "checkArrayForBubble", 1)
                sTmp = aMessageText(i)
                If InStr(sTmp,"post_bubble_middle") > 0 Then
                    sTmp = cutStrWithKey(sTmp, "post_bubble_middle", 18, "post_bubble_bottom", 0)
                    sTmp = cutStrWithKey(sTmp, ">", 1,   "</div>", 0)
                    aMessageText(i) = sTmp
                End If
            Next
        End Function

        Function saveOriginMsgTxt()
            Dim oTxt
            Set oTxt = Fso.OpenTextFile(uOriginMsgTxtPath, 8, False, True)

            For i = 0 To iOriginDataArrayLength
                oTxt.WriteLine(aPostNum(i))
                oTxt.WriteLine(aUserText(i))
                oTxt.WriteLine(aMessageText(i))
                oTxt.WriteLine()
            Next

            oTxt.Close
            Set oTxt = Nothing
        End Function





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



'*************************************************
'****get result.
'*************************************************
Function getPureResult(sMessageText, sUserText, iPostNum, sOptName)
    Dim aPickedResult, iNeedNum, aPureResult, isValid
    aPickedResult = pickMachedResults(sMessageText, sOptName)
    If isArray(aPickedResult) Then
        iNeedNum = getNeedNumOfResult(sOptName)
        isValid = checkResultsNum(aPickedResult, iNeedNum)
        getPureResult = Array(joinArrayWithSpace(aPickedResult), isValid)
    Else
        getPureResult = Array("", False)
    End If
End Function

        Function getNeedNumOfResult(sOptName)
            If (sOptName = "333bf") Or _
                    (sOptName = "333fm") Or _
                    (sOptName = "666") Or _
                    (sOptName = "777") Then
                getNeedNumOfResult = 3
            Else
                getNeedNumOfResult = 5
            End If
        End Function

        Function pickMachedResults(sMessageText, sOptName)
            Dim aTmpResult, sTmp, sTmpResult
            'MsgBox("sMessageText="&sMessageText&Vblf&_
                    '"replaceCharacterInResultStr(sMessageText)="&replaceCharacterInResultStr(sMessageText))
            aTmpResult = Split(replaceCharacterInResultStr(cutResultStrByKeyword(sMessageText)))
            For i = 0 To safeUBound(aTmpResult, "pickMachedResults", 1)
                If sOptName = "333fm" Then
                    sTmp = formatResultStrForFm(aTmpResult(i))
                Else
                    sTmp = formatResultStr(aTmpResult(i))
                End If

                If IsNumeric(sTmp) Then
                    sTmpResult = sTmpResult & " " & sTmp
                End If
            Next
            'MsgBox(sTmpResult)

            pickMachedResults = Split(Trim(sTmpResult))
        End Function

        Function cutResultStrByKeyword(sMessageText)
            Dim iKeywordInStr, aKeyword, sTmp
            aKeyword = Array("成绩列表","详细时间")
            sTmp = sMessageText
            For i = 0 To safeUBound(aKeyword, "cutResultStrByKeyword", 1)
                iKeywordInStr = InStr(sMessageText, aKeyword(i))
                If iKeywordInStr > 0 Then
                    sTmp = safeMid(sMessageText, iKeywordInStr+4, 0, "cutResultStrByKeyword")
                    Exit For
                End If
            Next
            cutResultStrByKeyword = sTmp
        End Function

        Function checkResultsNum(aPickedResult, iNeedNum)
            Dim isValid
            If safeUBound(aPickedResult, "checkResultsNum", 1) + 1 = iNeedNum Then
                isValid = True
            Else
                isValid = False
            End If
            checkResultsNum = isValid
        End Function



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



'*************************************************
'****get avg and sorted results.
'*************************************************
Function getAvgAndSortedResults(sOptName, sPureResult, isValid)
    If isValid Then
        Dim aTmp, aPureResults, lenResults, iAvgResult, iEnd
        aPureResults = Split(sPureResult)
        aTmp = sortIntArray(aPureResults)
        lenResults = safeUBound(aTmp, "getAvgAndSortedResults", 1)

        If lenResults = 4 Then
            iEnd = 3
        Else
            iEnd = 2
        End If

        If aTmp(iEnd) <> 9999.99 Then
            iAvgResult = Round(((0 + aTmp(iEnd-2) + aTmp(iEnd-1) + aTmp(iEnd)) / 3), 2)
            iAvgResult = FormatNumber(iAvgResult, 2, , , 0)
        Else
            iAvgResult = 9999.99
        End If

        Dim sBrFlag
        sBrFlag = compareToRecord(sOptName, iAvgResult, aTmp(0))

        getAvgAndSortedResults = iAvgResult & " " & joinArrayWithSpace(aTmp) & " " & sBrFlag
    Else
        getAvgAndSortedResults = ""
    End If
End Function

Function loadAllRecords()
    Dim oTxt, count
    Set oTxt = Fso.OpenTextFile(uRecordTxt, 1)
    count = 0

    Do Until oTxt.AtEndOfStream
        aAllRecords(count) = oTxt.ReadLine
        count = count + 1
    Loop

    oTxt.Close
    Set oTxt = Nothing
End Function

Function compareToRecord(sOptName, iAvgResult, iBestResult)
    Dim iOptCount, aTmpRecord, iAvgRecord, iBestRecord, isAvgBr, isBestBr
    iOptCount = getSeqInAllOptName(sOptName)
    aTmpRecord = Split(aAllRecords(iOptCount))
    iAvgRecord = aTmpRecord(2)
    iBestRecord = aTmpRecord(1)
    isAvgBr = False
    isBestBr = False

    If sOptName <> "333fm" Then
        iAvgRecord = formatResultStr(iAvgRecord)
        iBestRecord = formatResultStr(iBestRecord)
    End If

    If iAvgResult - iAvgRecord <= 0 Then
        isAvgBr = True
    End If

    If iBestResult - iBestRecord <= 0 Then
        isBestBr = True
    End If

    compareToRecord = isAvgBr & " " & isBestBr
End Function



'*************************************************
'****sort all results.
'*************************************************
Function sortAllResults()
    sortByOptName()

    sortByAvgResult aTmpUserText_333, aTmpPureResults_333, aTmpSortedResults_333, 0
    sortByAvgResult aTmpUserText_444, aTmpPureResults_444, aTmpSortedResults_444, 0
    sortByAvgResult aTmpUserText_555, aTmpPureResults_555, aTmpSortedResults_555, 0
    sortByAvgResult aTmpUserText_222, aTmpPureResults_222, aTmpSortedResults_222, 0
    sortByAvgResult aTmpUserText_333bf, aTmpPureResults_333bf, aTmpSortedResults_333bf, 1
    sortByAvgResult aTmpUserText_333oh, aTmpPureResults_333oh, aTmpSortedResults_333oh, 0
    sortByAvgResult aTmpUserText_333fm, aTmpPureResults_333fm, aTmpSortedResults_333fm, 0
    sortByAvgResult aTmpUserText_py, aTmpPureResults_py, aTmpSortedResults_py, 0
    sortByAvgResult aTmpUserText_mega, aTmpPureResults_mega, aTmpSortedResults_mega, 0
    sortByAvgResult aTmpUserText_sq, aTmpPureResults_sq, aTmpSortedResults_sq, 0
    sortByAvgResult aTmpUserText_clock, aTmpPureResults_clock, aTmpSortedResults_clock, 0
    sortByAvgResult aTmpUserText_sk, aTmpPureResults_sk, aTmpSortedResults_sk, 0
    sortByAvgResult aTmpUserText_666, aTmpPureResults_666, aTmpSortedResults_666, 0
    sortByAvgResult aTmpUserText_777, aTmpPureResults_777, aTmpSortedResults_777, 0

    MsgBox("OK")
End Function

    Function sortByOptName()
        Dim aAllValidOptName, count, aTmpUserText(), aTmpPureResults(), aTmpSortedResults()
        aAllValidOptName = getAllOptName()
        count = 0
        

        For i = 0 To UBound(aAllValidOptName)
            ReDim Preserve aTmpUserText(0)
            ReDim Preserve aTmpPureResults(0)
            ReDim Preserve aTmpSortedResults(0)
            aTmpUserText(0) = ""
            aTmpPureResults(0) = ""
            aTmpSortedResults(0) = ""
            count = 0

            For j = 0 To iTmpResultCount - 1
                If aAllIsValid(j) And _
                        aAllOptName(j) = aAllValidOptName(i) And _
                        getExistInArray(aTmpUserText, aAllUserText(j)) = "" Then
                    ReDim Preserve aTmpUserText(count)
                    ReDim Preserve aTmpPureResults(count)
                    ReDim Preserve aTmpSortedResults(count)

                    aTmpUserText(count) = aAllUserText(j)
                    aTmpPureResults(count) = aAllPureResults(j)
                    aTmpSortedResults(count) = aAllSortedResults(j)
                    count = count + 1
                End If
            Next

            Select Case aAllValidOptName(i)
                Case "333"
                    aTmpUserText_333 = aTmpUserText
                    aTmpPureResults_333 = aTmpPureResults
                    aTmpSortedResults_333 = aTmpSortedResults
                Case "444"
                    aTmpUserText_444 = aTmpUserText
                    aTmpPureResults_444 = aTmpPureResults
                    aTmpSortedResults_444 = aTmpSortedResults
                Case "555"
                    aTmpUserText_555 = aTmpUserText
                    aTmpPureResults_555 = aTmpPureResults
                    aTmpSortedResults_555 = aTmpSortedResults
                Case "222"
                    aTmpUserText_222 = aTmpUserText
                    aTmpPureResults_222 = aTmpPureResults
                    aTmpSortedResults_222 = aTmpSortedResults
                Case "333bf"
                    aTmpUserText_333bf = aTmpUserText
                    aTmpPureResults_333bf = aTmpPureResults
                    aTmpSortedResults_333bf = aTmpSortedResults
                Case "333oh"
                    aTmpUserText_333oh = aTmpUserText
                    aTmpPureResults_333oh = aTmpPureResults
                    aTmpSortedResults_333oh = aTmpSortedResults
                Case "333fm"
                    aTmpUserText_333fm = aTmpUserText
                    aTmpPureResults_333fm = aTmpPureResults
                    aTmpSortedResults_333fm = aTmpSortedResults
                Case "py"
                    aTmpUserText_py = aTmpUserText
                    aTmpPureResults_py = aTmpPureResults
                    aTmpSortedResults_py = aTmpSortedResults
                Case "mega"
                    aTmpUserText_mega = aTmpUserText
                    aTmpPureResults_mega = aTmpPureResults
                    aTmpSortedResults_mega = aTmpSortedResults
                Case "sq"
                    aTmpUserText_sq = aTmpUserText
                    aTmpPureResults_sq = aTmpPureResults
                    aTmpSortedResults_sq = aTmpSortedResults
                Case "clock"
                    aTmpUserText_clock = aTmpUserText
                    aTmpPureResults_clock = aTmpPureResults
                    aTmpSortedResults_clock = aTmpSortedResults
                Case "sk"
                    aTmpUserText_sk = aTmpUserText
                    aTmpPureResults_sk = aTmpPureResults
                    aTmpSortedResults_sk = aTmpSortedResults
                Case "666"
                    aTmpUserText_666 = aTmpUserText
                    aTmpPureResults_666 = aTmpPureResults
                    aTmpSortedResults_666 = aTmpSortedResults
                Case "777"
                    aTmpUserText_777 = aTmpUserText
                    aTmpPureResults_777 = aTmpPureResults
                    aTmpSortedResults_777 = aTmpSortedResults
            End Select
        Next
    End Function

    Function sortByAvgResult(aTmpUserText, aTmpPureResults, aTmpSortedResults, iSeq)
        Dim lenTmpUserText, flag
        lenTmpUserText = safeUBound(aTmpUserText, "sortByAvgResult 111", 1)
        flag = 0

        If lenTmpUserText > 0 Then
            For i = 0 To lenTmpUserText - 1
                For j = i + 1 To lenTmpUserText
                    Dim aResults1, aResults2
                    aResults1 = Split(aTmpSortedResults(i))
                    aResults2 = Split(aTmpSortedResults(j))

                    If iSeq = 0 Then
                        flag = compareTwoResult(aResults1(iSeq), aResults2(iSeq), i, j, aTmpUserText, aTmpPureResults, aTmpSortedResults)
                    End If

                    If (iSeq = 1 Or flag = 1) Then
                        For k = 1 To safeUBound(aResults1, "sortByAvgResult 222", 1) - 2
                            flag = compareTwoResult(aResults1(k), aResults2(k), i, j, aTmpUserText, aTmpPureResults, aTmpSortedResults)
                            If flag = 0 Then
                                Exit For
                            End If
                        Next
                    End If
                Next
            Next
        End If
    End Function

        Function compareTwoResult(iResult1, iResult2, i, j, aTmpUserText, aTmpPureResults, aTmpSortedResults)
            Dim flag
            flag = 0
            'If Not IsNumeric(iResult1) Then
            '    MsgBox("iResult1=" & iResult1 & Vblf &_
            '            "aTmpUserText(i)=" & aTmpUserText(i) & Vblf &_
            '            "aTmpPureResults(i)=" & aTmpPureResults(i) & Vblf &_
            '            "aTmpSortedResults(i)=" & aTmpSortedResults(i))
            'End If
            If iResult1 - iResult2 > 0 Then
                swapTwoStrings aTmpUserText(i), aTmpUserText(j)
                swapTwoStrings aTmpPureResults(i), aTmpPureResults(j)
                swapTwoStrings aTmpSortedResults(i), aTmpSortedResults(j)
            ElseIf iResult1 - iResult2 = 0 Then
                flag = 1
            End If

            compareTwoResult = flag
        End Function



'*************************************************
'****write final result into excel file.
'*************************************************
Function writeReslutsToExcel()
    Dim folderPath, oldFilePath, newFilePath
    folderPath = ws.CurrentDirectory & "\allResults\"
    oldFilePath = folderPath & (iCompNum - 1) & ".xlsx"
    newFilePath = folderPath & iCompNum & ".xlsx"
    If Fso.FileExists(oldFilePath) And Not Fso.FileExists(newFilePath) Then
        Fso.CopyFile oldFilePath, newFilePath, False

        Dim ExcelApp, ExcelBook, ExcelSheet
        Set ExcelApp = CreateObject("Excel.Application")
        Set ExcelBook= ExcelApp.Workbooks.Open(newFilePath)

        ExcelApp.DisplayAlerts = False
        ExcelBook.Sheets(1).Delete
        ExcelApp.DisplayAlerts = True

        'Set ExcelSheet = ExcelBook.Sheets("Sheet1")
        Set ExcelSheet = ExcelBook.Sheets.Add

        ExcelSheet.Columns("B").ColumnWidth = 15

        writeAllProjectResult ExcelSheet, aTmpUserText_333, aTmpPureResults_333, aTmpSortedResults_333, "333"
        writeAllProjectResult ExcelSheet, aTmpUserText_444, aTmpPureResults_444, aTmpSortedResults_444, "444"
        writeAllProjectResult ExcelSheet, aTmpUserText_555, aTmpPureResults_555, aTmpSortedResults_555, "555"
        writeAllProjectResult ExcelSheet, aTmpUserText_222, aTmpPureResults_222, aTmpSortedResults_222, "222"
        writeAllProjectResult ExcelSheet, aTmpUserText_333bf, aTmpPureResults_333bf, aTmpSortedResults_333bf, "333bf"
        writeAllProjectResult ExcelSheet, aTmpUserText_333oh, aTmpPureResults_333oh, aTmpSortedResults_333oh, "333oh"
        writeAllProjectResult ExcelSheet, aTmpUserText_333fm, aTmpPureResults_333fm, aTmpSortedResults_333fm, "333fm"
        writeAllProjectResult ExcelSheet, aTmpUserText_py, aTmpPureResults_py, aTmpSortedResults_py, "py"
        writeAllProjectResult ExcelSheet, aTmpUserText_mega, aTmpPureResults_mega, aTmpSortedResults_mega, "mega"
        writeAllProjectResult ExcelSheet, aTmpUserText_sq, aTmpPureResults_sq, aTmpSortedResults_sq, "sq"
        writeAllProjectResult ExcelSheet, aTmpUserText_clock, aTmpPureResults_clock, aTmpSortedResults_clock, "clock"
        writeAllProjectResult ExcelSheet, aTmpUserText_sk, aTmpPureResults_sk, aTmpSortedResults_sk, "sk"
        writeAllProjectResult ExcelSheet, aTmpUserText_666, aTmpPureResults_666, aTmpSortedResults_666, "666"
        writeAllProjectResult ExcelSheet, aTmpUserText_777, aTmpPureResults_777, aTmpSortedResults_777, "777"
        
        ExcelBook.Save
        ExcelBook.Close
        ExcelApp.Quit
        Set ExcelSheet = Nothing
        Set ExcelBook = Nothing
        Set ExcelApp = Nothing

        If brResultCount > 0 Then
            writeBrToExcel newFilePath, True
        Else
            writeBrToExcel newFilePath, False
        End If
    Else
        Msgbox(oldFilePath & Vblf & "not exist" & Vblf & " or" & Vblf & newFilePath & Vblf & "already exist")
    End If

    Msgbox("Write done!")
End Function

        Function writeAllProjectResult(ExcelSheet, aUserText, aPureResults, aSortedResults, sOptName)
            Dim rankNum, isAvgBr, isBestBr
            rankNum = 0
            countExcelRow = countExcelRow + 1
            For i = 1 To 9
                If i = 1 Then
                    ExcelSheet.Cells(countExcelRow, i).HorizontalAlignment = 2
                ElseIf i = 2 Then
                    ExcelSheet.Cells(countExcelRow, i).HorizontalAlignment = 3
                Else
                    ExcelSheet.Cells(countExcelRow, i).HorizontalAlignment = 4
                End If
                ExcelSheet.Cells(countExcelRow, i).Interior.Color = RGB(0,0,0)
                ExcelSheet.Cells(countExcelRow, i).Font.Color = RGB(255,255,255)
            Next

            ExcelSheet.Cells(countExcelRow, 1).Value = getFinalOptName(sOptName)
            ExcelSheet.Cells(countExcelRow, 2).Value = "ID"
            ExcelSheet.Cells(countExcelRow, 3).Value = "最好成绩"
            ExcelSheet.Cells(countExcelRow, 4).Value = "平均成绩"
            ExcelSheet.Cells(countExcelRow, 5).Value = "r1"
            ExcelSheet.Cells(countExcelRow, 6).Value = "r2"
            ExcelSheet.Cells(countExcelRow, 7).Value = "r3"
            ExcelSheet.Cells(countExcelRow, 8).Value = "r4"
            ExcelSheet.Cells(countExcelRow, 9).Value = "r5"

            For i = 0 To safeUBound(aUserText, "writeAllProjectResult", 1)
                Dim aSplitSortedResults, aSplitPureResults
                countExcelRow = countExcelRow + 1
                rankNum = rankNum + 1
                aSplitSortedResults = Split(aSortedResults(i))
                aSplitPureResults = Split(aPureResults(i))
                lenTmpSortedResults = safeUBound(aSplitSortedResults, "writeAllProjectResult", 1)
                isAvgBr = aSplitSortedResults(lenTmpSortedResults - 1)
                isBestBr = aSplitSortedResults(lenTmpSortedResults)

                For j = 1 To 9
                    If j = 2 Then
                        ExcelSheet.Cells(countExcelRow, j).HorizontalAlignment = 2
                    Else
                        ExcelSheet.Cells(countExcelRow, j).HorizontalAlignment = 4
                    End If

                    If j = 4 Then
                        ExcelSheet.Cells(countExcelRow, j).Font.Bold = True
                    End If

                    If j < 3 Then
                        ExcelSheet.Cells(countExcelRow, j).Font.Color = RGB(247,83,9)
                    Else
                        ExcelSheet.Cells(countExcelRow, j).NumberFormat = "@"
                        ExcelSheet.Cells(countExcelRow, j).Font.Color = RGB(0,0,0)
                    End If

                    If isEvenNum(i) Then
                        ExcelSheet.Cells(countExcelRow, j).Interior.Color = RGB(255,255,255)
                    Else
                        ExcelSheet.Cells(countExcelRow, j).Interior.Color = RGB(230,230,230)
                    End If
                Next

                If isBestBr Then
                    ExcelSheet.Cells(countExcelRow, 3).Interior.Color = RGB(255,192,0)
                    saveBrResult sOptName, aUserText(i), aSplitSortedResults(1), False, aPureResults(i)
                End If
                If isAvgBr Then
                    ExcelSheet.Cells(countExcelRow, 4).Interior.Color = RGB(255,192,0)
                    saveBrResult sOptName, aUserText(i), aSplitSortedResults(0), True, aPureResults(i)
                End If

                ExcelSheet.Cells(countExcelRow, 1).Value = rankNum
                ExcelSheet.Cells(countExcelRow, 2).Value = aUserText(i)
                ExcelSheet.Cells(countExcelRow, 3).Value = revertResult(aSplitSortedResults(1), sOptName)
                ExcelSheet.Cells(countExcelRow, 4).Value = revertResult(aSplitSortedResults(0), sOptName)
                For j = 0 To UBound(aSplitPureResults)
                    ExcelSheet.Cells(countExcelRow, j+5).Value = revertResult(aSplitPureResults(j), sOptName)
                Next
            Next

            countExcelRow = countExcelRow + 1
        End Function

                Function saveBrResult(sOptName, sUserText, brResult, isAvg, sPureResult)
                    ReDim Preserve brResultData(4,brResultCount)
                    brResultData(0,brResultCount) = sOptName
                    brResultData(1,brResultCount) = sUserText
                    brResultData(2,brResultCount) = brResult
                    brResultData(3,brResultCount) = isAvg
                    brResultData(4,brResultCount) = sPureResult
                    brResultCount = brResultCount + 1
                End Function



'*************************************************
'****write BR into excel file.
'*************************************************
Function writeBrToExcel(newFilePath, haveNewBr)
    Dim ExcelApp, ExcelBook, ExcelSheet, brLine, aSplitPureResults
    Set ExcelApp = CreateObject("Excel.Application")
    Set ExcelBook= ExcelApp.Workbooks.Open(newFilePath)
    Set ExcelSheet = ExcelBook.Sheets("Sheet2")

    For i = 1 To 57
        For j = 1 To 10
            If i mod 4 = 0 Then
                ExcelSheet.Cells(i, j).Interior.Color = RGB(230,230,230)
            Else
                ExcelSheet.Cells(i, j).Interior.Color = RGB(250,250,250)
            End If
        Next
    Next
            
    If haveNewBr Then
        For i = 0 To brResultCount-1
            aSplitPureResults = Split(brResultData(4,i))
            brLine = getBrLine(brResultData(0,i), brResultData(3,i))

            For j = 1 To 10
                ExcelSheet.Cells(brLine, j).Interior.Color = RGB(255,218,101)
            Next

            ExcelSheet.Cells(brLine, 2).Value = brResultData(2,i)
            ExcelSheet.Cells(brLine, 4).Value = brResultData(1,i)
            ExcelSheet.Cells(brLine, 5).Value = iCompNum & "期"

            If brResultData(3,i) Then
                For j = 0 To UBound(aSplitPureResults)
                    ExcelSheet.Cells(brLine, j + 6).Value = aSplitPureResults(j)
                Next
            End If
        Next
    End If

    ExcelBook.Save
    ExcelBook.Close
    ExcelApp.Quit
    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelApp = Nothing
End Function

        Function getBrLine(sOptName, isAvg)
            Dim tmpLine
            tmpLine = getSeqInAllOptName(sOptName)
            tmpLine = (tmpLine + 1) * 4
            If isAvg Then
                tmpLine = tmpLine + 1
            End If
            getBrLine = tmpLine
        End Function