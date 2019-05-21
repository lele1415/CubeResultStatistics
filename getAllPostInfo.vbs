'*************************************************
'****get data array.
'*************************************************
Const ID_IGNORE_POST_NUM = "ignore_post_num"

Dim vaAllPostInfo : Set vaAllPostInfo = New VariableArray
Dim aHeadStr : aHeadStr = Array("post_no&quot;:", "alog-group=""p_author""", "j_d_post_content  clearfix")
Dim aEndStr : aEndStr = Array(",", "</a>", "</div><br>            </cc>")

Const GET_POST_NUM = 0
Const GET_POST_USER = 1
Const GET_POST_MSG = 2

Dim iSeq : iSeq = GET_POST_NUM
Dim oTmp
Dim aIgnoredPostNum

Const ID_GET_RESULT_STATUS = "get_result_status"

Sub onClickGetAllPostInfo()
    Call setInnerHtml(ID_GET_RESULT_STATUS, "获取回复信息...")
    idTimer = window.setTimeout("getAllPostInfo()", 10, "VBScript")
End Sub

Sub getAllPostInfo()
    window.clearTimeout(idTimer)

    Call vaAllPostInfo.ResetArray()
    '//these post nums will be ignored
    aIgnoredPostNum = Split(getElementValue(ID_IGNORE_POST_NUM), " ")

    '//read text of page code
    Dim oTxt, sReadLine
    Set oTxt = Fso.OpenTextFile(uPagesCodeFile, 1, False, True)

    Do Until oTxt.AtEndOfStream
        sReadLine = oTxt.ReadLine

        If iSeq <> GET_POST_NUM Then
            Call searchInfo(sReadLine)
            '//post num maybe in the same line with prev post msg
            If iSeq = GET_POST_NUM Then Call searchInfo(sReadLine)
        Else
            Call searchInfo(sReadLine)
        End If
    Loop

    Call oTxt.Close
    Set oTxt = Nothing

    Call saveOriginMsgTxt()

    'MsgBox("getAllPostInfo done!")
    Call getAllResultInfo()
End Sub

        Sub searchInfo(sOrigin)
            Dim sGet
            Select Case iSeq
                '//get post num
                Case GET_POST_NUM
                    sGet = cutStrWithHeadEndStr(sOrigin, aHeadStr(0), aEndStr(0))
                    If sGet = "" Then Exit Sub

                    If Not checkIsValidPostNum(sGet) Then Exit Sub

                    '//new PostInfo
                    Set oTmp = New PostInfo
                    oTmp.PostNum = sGet
                '//get post user
                Case GET_POST_USER
                    sGet = cutStrWithElement(sOrigin, aHeadStr(iSeq), aEndStr(iSeq))
                    If sGet = "" Then Exit Sub

                    Call removeElement(sGet, "<img")
                    If Not checkIsValidResultOwner(sGet) Then Exit Sub

                    oTmp.PostUser = sGet
                '//get post msg
                Case GET_POST_MSG
                    sGet = cutStrWithElement(sOrigin, aHeadStr(iSeq), aEndStr(iSeq))
                    If sGet = "" Then Exit Sub

                    sGet = LTrim(sGet)
                    Call checkArrayForBubble(sGet)
                    Call removeHtmlStr(sGet)

                    oTmp.PostMsg = sGet
            End Select


            '//already get a PostInfo
            If iSeq = GET_POST_MSG Then
                Dim oNew : Set oNew = oTmp
                vaAllPostInfo.Append(oNew)
                Set oTmp = Nothing
            End If
            
            Call seqPlus()
        End Sub

        Function checkIsValidPostNum(num)
            Dim i
            For i = 0 To UBound(aIgnoredPostNum)
                If num = aIgnoredPostNum(i) Then
                    checkIsValidPostNum = False
                    Exit Function
                End If
            Next
            checkIsValidPostNum = True
        End Function

        Function checkIsValidResultOwner(name)
            If vaIgnoreUsers.Bound > -1 Then
                Dim i
                For i = 0 To vaIgnoreUsers.Bound
                    If StrComp(name, vaIgnoreUsers.V(i)) = 0 Then
                        checkIsValidResultOwner = False
                        Exit Function
                    End If
                Next
            End If
            checkIsValidResultOwner = True
        End Function

        Sub seqPlus()
            iSeq = iSeq + 1
            If iSeq = 3 Then iSeq = GET_POST_NUM
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
