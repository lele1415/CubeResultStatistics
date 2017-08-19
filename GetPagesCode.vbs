
'*************************************************
'****get code from Pages, and write into text file.
'*************************************************
Sub getPagesCode()
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
    Msgbox("Done!")
End Sub

        Sub receiveCode(sCode)
            Call oTxtPagesCode.Write(sCode)
            Call setInnerHtml("PageNum_id", iCrtPageNum)

            If iMaxPageNum = 0 Then
                iMaxPageNum = cutStrWithHeadEndStr(sCode, "total_page"":", "};")
            End If

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
