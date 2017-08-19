'*************************************************
'****get data array.
'*************************************************
Function getMsgArrayFromCode()
    Dim bContinue
    aPostNum = getArrayFromTxt("post_no&quot", "post_no&quot", 14, ",", -1)
    aUserText = getArrayFromTxt("alog-group=""p_author""", "_blank", 8, "</a>", -1)
    aMessageText = getArrayFromTxt("j_d_post_content", "j_d_post_content", 40, "user-hide-post-down", -32)
    iCompNum = getStrInTextFile(uCodeTxtPath, "<title>", "第", "期")
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
                    sTmp = cutStrWithHeadEndStr(sTmp, "post_bubble_middle", "post_bubble_bottom")
                    sTmp = cutStrWithHeadEndStr(sTmp, ">", "</div>")
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
