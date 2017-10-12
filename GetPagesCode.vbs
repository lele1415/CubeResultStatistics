
'*************************************************
'****get code from Pages, and write into text file.
'*************************************************
Const ID_URL = "url"
Const ID_GET_PAGE_CODE_STATUS = "get_page_code_status"
Const ID_CRT_PAGE_NUM = "crt_page_num"
Const ID_MAX_PAGE_NUM = "max_page_num"

Dim iCrtPageNum : iCrtPageNum = 1
Dim iMaxPageNum : iMaxPageNum = 0
Dim iCompNum : iCompNum = 0
Dim oTxtPagesCode
Dim bGetCodeDone : bGetCodeDone = False
Dim rootUrl

Sub getPagesCode()
    Call setInnerHtml(ID_GET_PAGE_CODE_STATUS, "正在获取...")
    rootUrl = getElementValue(ID_URL)

    Call initTxtFile(uPagesCodeFile)
    Set oTxtPagesCode = Fso.OpenTextFile(uPagesCodeFile, 8, False, True)

    '//get first page code
    Call getSourceJs(rootUrl)
    Call waitForGetCode()

    If Not checkMaxPageNum() Then Exit Sub

    '//get other pages code
    Dim i
    For i = 2 To iMaxPageNum
        Call getNextPageCode()
        Call waitForGetCode()
    Next
    
    Call oTxtPagesCode.Close
    Set oTxtPagesCode = Nothing
    Call checkCompNum()
    Call setInnerHtml(ID_GET_PAGE_CODE_STATUS, "已完成！")
End Sub

        Sub receiveCode(sCode)
            Call oTxtPagesCode.Write(sCode)

            If iMaxPageNum = 0 Then
                iMaxPageNum = cutStrWithHeadEndStr(sCode, "total_page"":", "};")
                Dim sTitle : sTitle = cutStrWithHeadEndStr(sCode, "<title>", "</title>")
                iCompNum = cutStrWithHeadEndStr(sTitle, "第", "期")
            End If

            Call setInnerHtml(ID_CRT_PAGE_NUM, iCrtPageNum)
            Call setInnerHtml(ID_MAX_PAGE_NUM, iMaxPageNum)

            bGetCodeDone = True
        End Sub

        Sub getNextPageCode()
            bGetCodeDone = False
            iCrtPageNum = iCrtPageNum + 1

            Call getNextSourceJs(rootUrl & "?pn=" & iCrtPageNum)
        End Sub

        Sub waitForGetCode()
            Do While True
                If bGetCodeDone Then
                    Exit Do
                Else
                    Sleep(100)
                End If
            Loop
        End Sub

        Function checkMaxPageNum()
            If Not (IsNumeric(iMaxPageNum)) Then
                Msgbox("iMaxPageNum is not number! iMaxPageNum=" & iMaxPageNum)
                checkMaxPageNum = False
                Exit Function
            End If

            If iMaxPageNum < 2 Then
                Msgbox("iMaxPageNum < 2 iMaxPageNum=" & iMaxPageNum)
                checkMaxPageNum = False
                Exit Function
            End If

            checkMaxPageNum = True
        End Function

        Sub checkCompNum()
            If Not IsNumeric(iCompNum) Then
                iCompNum = InputBox("Input CompNum")
                Do Until IsNumeric(iCompNum)
                    iCompNum = InputBox("Input CompNum")
                Loop
            End If
        End Sub
