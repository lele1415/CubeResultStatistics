
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
    iMaxPageNum = getStrInTextFile(uCodeTxtPath, "total_page", "total_page"":", "};")
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
            If Fso.FileExists(uSleepVbsPath)=False Then 
                Set objOutputFile = fso.CreateTextFile(uSleepVbsPath, True) 
                objOutputFile.Write "wscript.sleep WScript.Arguments(0)" 
                objOutputFile.Close 
            End If 
             CreateObject("WScript.Shell").Run uSleepVbsPath & " " & MSecs, 1 , True 
        End Sub

        Function getCode(url)
            Dim xmlHTTP
            Set xmlHTTP = CreateObject("MSXML2.ServerXMLHTTP")
            xmlHTTP.open "GET", url
            xmlHttp.send
            writePageCodeToTxt(xmlHttp.responseText)
            Set xmlHTTP = Nothing
        End Function

        Function writePageCodeToTxt(responseText)
            Set oTxt = Fso.OpenTextFile(uCodeTxtPath, 8, False, True)
            oTxt.Write(responseText)
            oTxt.Close
            Set oTxt = Nothing
        End Function
