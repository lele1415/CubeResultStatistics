'*************************************************
'****get data array.
'*************************************************
Dim vaAllPostInfo : Set vaAllPostInfo = New VariableArray
Dim aHeadStr : aHeadStr = Array("post_no&quot;:", "alog-group=""p_author""", "j_d_post_content  clearfix")
Dim aEndStr : aEndStr = Array(",", "</a>", "</div><br></cc>")

Dim iSeq : iSeq = 0
Dim oTmp

Sub getDetailFromCode()
    Dim oTxt, sReadLine
    Set oTxt = Fso.OpenTextFile(uPagesCodeFile, 1, False, True)

    Do Until oTxt.AtEndOfStream
        sReadLine = oTxt.ReadLine

        If iSeq > 0 Then
            Call searchInfo(sReadLine)
            If iSeq = 0 Then Call searchInfo(sReadLine)
        Else
            Call searchInfo(sReadLine)
        End If
    Loop

    Call oTxt.Close
    Set oTxt = Nothing

    Call saveOriginMsgTxt()

    MsgBox("getDetailFromCode done!")
    'spiltStrAndGetOptInfo()
End Sub

        Sub searchInfo(sOrigin)
            Dim sGet
            If iSeq = 0 Then
                sGet = cutStrWithHeadEndStr(sOrigin, aHeadStr(0), aEndStr(0))
            Else
                sGet = cutStrWithElement(sOrigin, aHeadStr(iSeq), aEndStr(iSeq))
            End If
                
            If sGet <> "" Then
                If iSeq = 0 Then Set oTmp = New PostInfo

                Call receiveInfoStr(sGet, oTmp)

                If iSeq = 2 Then
                    Dim oNew : Set oNew = oTmp
                    vaAllPostInfo.Append(oNew)
                    Set oTmp = Nothing
                End If
                
                Call seqPlus()
            End If
        End Sub

        Sub seqPlus()
            iSeq = iSeq + 1
            If iSeq = 3 Then iSeq = 0
        End Sub

        Sub receiveInfoStr(str, object)
            Select Case iSeq
                Case 0
                    object.PostNum = str
                Case 1
                    object.UserID = str
                Case 2
                    str = Mid(str, 13)
                    Call checkArrayForBubble(str)
                    Call removeHtmlStr(str)
                    object.MsgStr = str
            End Select
        End Sub

        Sub checkArrayForBubble(sMsg)
            If InStr(sMsg, "post_bubble_middle") > 0 Then
                Dim sTmp
                sTmp = cutStrWithElementRev(sMsg, "post_bubble_middle", "post_bubble_bottom")
                If sTmp <> "" Then sMsg = sTmp
            End If
        End Sub

        Sub saveOriginMsgTxt()
            initTxtFile(uOriginMsgTxtPath)
            Dim oTxt, i
            Set oTxt = Fso.OpenTextFile(uOriginMsgTxtPath, 8, False, True)

            For i = 0 To vaAllPostInfo.Length
                oTxt.WriteLine(vaAllPostInfo.V(i).PostNum)
                oTxt.WriteLine(vaAllPostInfo.V(i).UserID)
                oTxt.WriteLine(vaAllPostInfo.V(i).MsgStr)
                oTxt.WriteLine()
            Next

            oTxt.Close
            Set oTxt = Nothing
        End Sub
