'*************************************************
'****get data array.
'*************************************************
Dim vaAllPostInfo : Set vaAllPostInfo = New VariableArray
Dim aHeadStr : aHeadStr = Array("post_no&quot;:", "alog-group=""p_author""", "j_d_post_content  clearfix")
Dim aEndStr : aEndStr = Array(",", "</a>", "</div><br></cc>")

Dim iSeq : iSeq = 0
Dim oTmp
Dim aIgnoredPostNum

Sub getAllPostInfo()
    Call vaAllPostInfo.ResetArray()
    '//these post nums will be ignored
    aIgnoredPostNum = Split(getElementValue("exceptPostNum_id"), " ")

    '//read text of page code
    Dim oTxt, sReadLine
    Set oTxt = Fso.OpenTextFile(uPagesCodeFile, 1, False, True)

    Do Until oTxt.AtEndOfStream
        sReadLine = oTxt.ReadLine

        If iSeq > 0 Then
            Call searchInfo(sReadLine)
            '//post num maybe in the same line with last post msg
            If iSeq = 0 Then Call searchInfo(sReadLine)
        Else
            Call searchInfo(sReadLine)
        End If
    Loop

    Call oTxt.Close
    Set oTxt = Nothing

    Call saveOriginMsgTxt()

    MsgBox("getAllPostInfo done!")
    Call getAllResultInfo()
End Sub

        Sub searchInfo(sOrigin)
            Dim sGet
            '//get post num
            If iSeq = 0 Then
                sGet = cutStrWithHeadEndStr(sOrigin, aHeadStr(0), aEndStr(0))
                If Not checkIsNeedPostNum(sGet) Then Exit Sub
            '//get post user or msg
            Else
                sGet = cutStrWithElement(sOrigin, aHeadStr(iSeq), aEndStr(iSeq))
            End If
                
            If sGet <> "" Then
                '//new PostInfo
                If iSeq = 0 Then Set oTmp = New PostInfo

                Call receiveInfoStr(sGet, oTmp)

                '//already get a PostInfo
                If iSeq = 2 Then
                    Dim oNew : Set oNew = oTmp
                    vaAllPostInfo.Append(oNew)
                    Set oTmp = Nothing
                End If
                
                Call seqPlus()
            End If
        End Sub

        Function checkIsNeedPostNum(num)
            Dim i
            For i = 0 To UBound(aIgnoredPostNum)
                If num = aIgnoredPostNum(i) Then
                    checkIsNeedPostNum = False
                    Exit Function
                End If
            Next
            checkIsNeedPostNum = True
        End Function

        Sub seqPlus()
            iSeq = iSeq + 1
            If iSeq = 3 Then iSeq = 0
        End Sub

        Sub receiveInfoStr(str, object)
            Select Case iSeq
                Case 0
                    object.PostNum = str
                Case 1
                    Call removeElement(str, "<img")
                    object.PostUser = str
                Case 2
                    str = LTrim(str)
                    Call checkArrayForBubble(str)
                    Call removeHtmlStr(str)
                    object.PostMsg = str
            End Select
        End Sub

        Sub checkArrayForBubble(sMsg)
            If InStr(sMsg, "post_bubble_middle") > 0 Then
                Dim sTmp
                sTmp = cutStrWithElement(sMsg, "post_bubble_middle", "</div>")
                If sTmp <> "" Then sMsg = sTmp
            End If
        End Sub

        Sub saveOriginMsgTxt()
            initTxtFile(uOriginMsgTxtPath)
            Dim oTxt, i
            Set oTxt = Fso.OpenTextFile(uOriginMsgTxtPath, 8, False, True)

            For i = 0 To vaAllPostInfo.Bound
                oTxt.WriteLine(vaAllPostInfo.V(i).PostNum)
                oTxt.WriteLine(vaAllPostInfo.V(i).PostUser)
                oTxt.WriteLine(vaAllPostInfo.V(i).PostMsg)
                oTxt.WriteLine()
            Next

            oTxt.Close
            Set oTxt = Nothing
        End Sub
