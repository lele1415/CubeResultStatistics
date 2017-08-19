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
    sOrigin = RePlace(sOrigin, "<br>", " ")
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
                sSimple = sSimple & "D"
            Case InStr("nN", sChar) > 0
                sSimple = sSimple & "N"
            Case InStr("fF", sChar) > 0
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
        iColonCount = getCharCountInStr(sTmp, ":")
        iPointCount = getCharCountInStr(sTmp, ".")
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

Function replaceStrForOptName(sStr)
    Dim sTmp
    sTmp = sStr
    sTmp = RePlace(sTmp, "[", "")
    sTmp = RePlace(sTmp, "]", "")
    sTmp = LCase(sTmp)
    replaceStrForOptName = sTmp
End Function