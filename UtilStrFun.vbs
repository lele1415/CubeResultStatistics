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

Function swapTwoStrings(s1, s2)
    Dim sTmp
    sTmp = s1
    s1 = s2
    s2 = sTmp
End Function

Function getStrInTextFile(path, sKey, sHead, sEnd)
    Dim oTxt, sLine, sGet
    Set oTxt = Fso.OpenTextFile(path, 1, False, True)

    Do Until oTxt.AtEndOfStream
        sLine = oTxt.ReadLine
        If InStr(sLine, sKey) > 0 Then
            sGet = cutStrWithHeadEndStr(sLine, sHead, sEnd)
            Exit Do
        End If
    Loop
    oTxt.Close
    Set oTxt = Nothing
    getStrInTextFile = sGet
End Function

Function cutStrWithHeadEndStr(sOrigin, sHead, sEnd)
    Dim iHead
    iHead = InStr(sOrigin, sHead) + Len(sHead)
    If iHead = Len(sHead) Then cutStrWithHeadEndStr = "" : Exit Function

    Dim iEnd
    iEnd = InStr(iHead, sOrigin, sEnd)
    If iEnd = Len(sHead) Then cutStrWithHeadEndStr = "" : Exit Function

    cutStrWithHeadEndStr = Mid(sOrigin, iHead, iEnd - iHead)
End Function

Function cutStrWithElement(sOrigin, sHead, sEnd)
    Dim iHead
    iHead = InStr(sOrigin, sHead)
    If iHead = 0 Then cutStrWithElement = "" : Exit Function

    iHead = InStr(iHead + Len(sHead), sOrigin, ">") + 1
    If iHead = 1 Then cutStrWithElement = "" : Exit Function

    Dim iEnd
    iEnd = InStr(iHead, sOrigin, sEnd)
    If iEnd = 0 Then cutStrWithElement = "" : Exit Function

    cutStrWithElement = Mid(sOrigin, iHead, iEnd - iHead)
End Function

Sub removeHtmlStr(sOrigin)
    Dim iImgStart, iImgLen, iDivStart, iDivLen
    sOrigin = RePlace(sOrigin, "<br>", VbCrlf)
    sOrigin = RePlace(sOrigin, "&#39;", "'")
    sOrigin = RePlace(sOrigin, "</a>", "")
    sOrigin = RePlace(sOrigin, "</div>", "")
    sOrigin = RePlace(sOrigin, "</span>", "")
    
    Call removeElement(sOrigin, "<img")
    Call removeElement(sOrigin, "<div")
    Call removeElement(sOrigin, "<embed")
    Call removeElement(sOrigin, "<a href")

    Do While (InStr(sOrigin, "  ") > 0)
        sOrigin = RePlace(sOrigin, "  ", " ")
    Loop
End Sub

Sub removeElement(sOrigin, sElement)
    Dim iHead, iEnd
    iHead = InStr(sOrigin, sElement)
    Do While (iHead > 0)
        iEnd = InStr(iHead, sOrigin, ">")
        If iEnd = 0 Then Exit Do
        sOrigin = RePlace(sOrigin, Mid(sOrigin, iHead, iEnd - iHead + 1), "")
        iHead = InStr(sOrigin, sElement)
    Loop
End Sub

Sub replaceCharForOptName(str)
    Dim sTmp, sChar, i
    sTmp = ""
    For i = 1 To Len(str)
        sChar = Mid(str, i, 1)
        Select Case sChar
            Case "二"
                sTmp = sTmp & "2"
            Case "三"
                sTmp = sTmp & "3"
            Case "四"
                sTmp = sTmp & "4"
            Case "五"
                sTmp = sTmp & "5"
            Case "六"
                sTmp = sTmp & "6"
            Case "七"
                sTmp = sTmp & "7"
            Case Else
                sTmp = sTmp & sChar
        End Select
    Next
    str = LCase(sTmp)
End Sub

Function isEvenNum(iNum)
    Dim numTmp
    numTmp = iNum / 2
    
    If numTmp = Round(numTmp, 0) Then
        isEvenNum = True
        Exit Function
    End If

    isEvenNum = False
End Function

Function getCharCountInStr(sStr, sChar)
    Dim iStep, sTmp, iCount
    sTmp = sStr
    iCount = 0
    Do While (InStr(sTmp, sChar) > 0)
        sTmp = RePlace(sTmp, sChar, "", 1, 1)
        iCount = iCount + 1
    Loop

    getCharCountInStr = iCount
End Function

Function replaceCharacterInResultStr(sStr)
    Dim sTmp, sSimple, sChar
    sTmp = sStr
    sSimple = ""
    For searchStep = 1 To Len(sTmp)
        sChar = safeMid(sTmp, searchStep, 1, "replaceCharacterInResultStr")
        Select Case True
            Case IsNumeric(sChar)
                sSimple = sSimple & sChar
            Case InStr(".。", sChar) > 0
                sSimple = sSimple & "."
            Case InStr(":：", sChar) > 0
                sSimple = sSimple & ":"
            Case InStr(",，", sChar) > 0
                sSimple = sSimple & " "
            Case InStr("dD", sChar) > 0
                sSimple = sSimple & "d"
            Case InStr("nN", sChar) > 0
                sSimple = sSimple & "n"
            Case InStr("fF", sChar) > 0
                sSimple = sSimple & "f"
            Case InStr("sS", sChar) > 0
                sSimple = sSimple & "s"
            Case InStr("/", sChar) > 0
                sSimple = sSimple & "/"
            Case Else
                sSimple = sSimple & " "
        End Select
    Next

    Do Until Not (InStr(sSimple, "  ") > 0)
        sSimple = RePlace(sSimple, "  ", " ")
    Loop

    replaceCharacterInResultStr = Trim(sSimple)
End Function

Function formatBrStr(sStr)
    Dim sPartMin, sPartSec, aTmp
    aTmp = Split(sStr, ":")
    sPartMin = aTmp(0)

    If Not IsNumeric(sPartMin) Then
        Exit Function
    End If

    sPartSec = aTmp(1)

    sStr = sPartMin * 60 + sPartSec
    formatBrStr = FormatNumber(sStr, 2, , , 0)
End Function

Sub formatResultStr(sStr, optSeq, vaObj)
    If StrComp(optSeq, OPT_SEQ_3fm) = 0 Then
        Call formatResultStrForFm(sStr, vaObj)
    ElseIf StrComp(optSeq, OPT_SEQ_3mb) = 0 Then
        Call formatResultStrFor3mb(sStr, vaObj)
    Else
        Call formatResultStrForOthers(sStr, vaObj)
    End If
End Sub

Sub formatResultStrForOthers(sStr, vaObj)
    '//DNF or DNS
    If InStr(sStr, "dnf") > 0 Or InStr(sStr, "dns") > 0 Then
        Call vaObj.Append(9999.99)
        Exit Sub
    End If

    sStr = RePlace(sStr, "d", "")
    sStr = RePlace(sStr, "n", "")
    sStr = RePlace(sStr, "f", "")

    Dim iColonCount, iPointCount
    iColonCount = getCharCountInStr(sStr, ":")
    iPointCount = getCharCountInStr(sStr, ".")

    '//invalid result str of xx:xx:xx.xx ???:xx
    If (iColonCount > 1) Or (iPointCount <> 1) Then
        Exit Sub
    End If

    Dim iLenOfStr, iPointInStr
    iLenOfStr = Len(sStr)
    iPointInStr = InStr(sStr, ".")

    '//invalid result str of .xx ???.x ???.
    If iPointInStr = 1 Or (iLenOfStr - iPointInStr < 2) Then
        Exit Sub
    End If

    '//if have colon
    If iColonCount = 1 Then
        Dim iColonInStr
        iColonInStr = InStr(sStr, ":")

        '//invalid result str of xx:.xx  :xx  ???:
        If (iPointInStr - iColonInStr) < 2 Or _
                iColonInStr = 1 Or _
                (iLenOfStr - iColonInStr) < 4 Then
            Exit Sub
        End If
    End If

    '//cut ???.xx
    Dim sFormat
    sFormat = safeMid(sStr, 1, iPointInStr + 2, "formatResultStrForOthers 111")

    '//get final str
    If iColonCount = 1 Then
        Dim sPartMin, sPartSec, aTmp
        aTmp = Split(sFormat, ":")
        sPartMin = aTmp(0)

        If Not IsNumeric(sPartMin) Then
            Exit Sub
        End If

        sPartSec = aTmp(1)

        sFormat = sPartMin * 60 + sPartSec
        sFormat = FormatNumber(sFormat, 2, , , 0)
    Else
        sFormat = FormatNumber(sFormat, 2, , , 0)
    End If

    Call vaObj.Append(sFormat)
End Sub

Sub formatResultStrForFm(sStr, vaObj)
    '//DNF or DNS
    If InStr(sStr, "dnf") > 0 Or InStr(sStr, "dns") > 0 Then
        Call vaObj.Append(9999.99)
        Exit Sub
    End If

    sStr = RePlace(sStr, "d", "")
    sStr = RePlace(sStr, "n", "")
    sStr = RePlace(sStr, "f", "")

    If Not IsNumeric(sStr) Then
        Exit Sub
    End If

    If Len(sStr) < 2 Then
        Exit Sub
    End If

    Dim iPointInStr, iColonInStr
    iPointInStr = InStr(sStr, ".")
    iColonInStr = InStr(sStr, ":")

    If iPointInStr > 0 Or iColonInStr > 0 Then
        Exit Sub
    End If

    Call vaObj.Append(sStr)
End Sub

Sub formatResultStrFor3mb(sStr, vaObj)
    '//DNF or DNS
    If InStr(sStr, "dnf") > 0 Or InStr(sStr, "dns") > 0 Then
        Call vaObj.Append(9999.99)
        Exit Sub
    End If

    sStr = RePlace(sStr, "d", "")
    sStr = RePlace(sStr, "n", "")
    sStr = RePlace(sStr, "f", "")

    Dim iSlashInStr
    iSlashInStr = InStr(sStr, "/")
    If iSlashInStr > 0 Then
        If iSlashInStr = 1 Or Len(sStr) = iSlashInStr Then Exit Sub
        Dim aTmp1
        aTmp1 = Split(sStr, "/")
        If UBound(aTmp1) <> 1 Then Exit Sub
        If (Not IsNumeric(aTmp1(0))) Or (Not IsNumeric(aTmp1(1))) Then Exit Sub
        Call vaObj.Append(aTmp1(0))
        Call vaObj.Append(aTmp1(1))
        Exit Sub
    End If

    Dim iColonCount
    iColonCount = getCharCountInStr(sStr, ":")

    '//invalid result str of xx:xx:xx.xx
    If (iColonCount > 1) Then
        Exit Sub
    End If

    Dim iLenOfStr, iPointInStr, iColonInStr
    iLenOfStr = Len(sStr)
    iPointInStr = InStr(sStr, ".")
    iColonInStr = InStr(sStr, ":")

    '//cut ???:xx
    If iPointInStr > 0 Then
        sStr = safeMid(sStr, 1, iPointInStr - 1, "formatResultStrFor3mb 111")
    End If

    '//if have colon
    If iColonCount = 1 Then

        '//invalid result str of xx:.xx  :xx  ???:
        If iColonInStr = 1 Or _
                (iLenOfStr - iColonInStr) <> 2 Then
            Exit Sub
        End If
    End If

    '//get final str
    If iColonCount = 1 Then
        Dim sPartMin, sPartSec, aTmp2
        aTmp2 = Split(sStr, ":")
        sPartMin = aTmp2(0)

        If Not IsNumeric(sPartMin) Then
            Exit Sub
        End If

        sPartSec = aTmp2(1)

        sStr = sPartMin * 60 + sPartSec
    End If

    Call vaObj.Append(sStr)
End Sub

Function revertResult(iResult, optSeq)
    Dim sRevert
    sRevert = iResult

    If sRevert = 9999.99 Then
        revertResult = "DNF"
        Exit Function
    End If

    If StrComp(optSeq, OPT_SEQ_3fm) <> 0 Then
        Dim minNum, secNum
        If sRevert > 59.99 Then
            minNum = int(sRevert / 60)

            If StrComp(optSeq, OPT_SEQ_3mb) <> 0 Then
                secNum = FormatNumber(sRevert - minNum * 60, 2, , , 0)
            Else
                secNum = FormatNumber(sRevert - minNum * 60, 0, , , 0)
                If secNum = "" Then secNum = 0
            End If

            If secNum = 0 Then
                secNum = "00"
            ElseIf secNum < 1.00 Then
                secNum = "00" & secNum
            ElseIf secNum < 10.00 Then
                secNum = "0" & secNum
            End If
            sRevert = minNum & ":" & secNum
        Else
            If StrComp(optSeq, OPT_SEQ_3mb) <> 0 Then
                sRevert = FormatNumber(sRevert, 2, , , 0)
                If sRevert < 1.00 Then
                    sRevert = "0" & sRevert
                End If
            Else
                sRevert = FormatNumber(sRevert, 0, , , 0)
            End If
        End If
    End If
        
    revertResult = sRevert
End Function

Function replaceStrForOptName(sStr)
    Dim sTmp
    sTmp = sStr
    sTmp = RePlace(sTmp, "[", "")
    sTmp = RePlace(sTmp, "]", "")
    sTmp = LCase(sTmp)
    replaceStrForOptName = sTmp
End Function
