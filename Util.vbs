Dim width,height
width=CreateObject("HtmlFile").ParentWindow.Screen.AvailWidth
Window.MoveTo width-700,100
Window.ResizeTo 700,700

Dim ws, Fso
Set ws=CreateObject("wscript.shell")
Set Fso=CreateObject("Scripting.FileSystemObject")

Dim CrtPath : CrtPath = ws.CurrentDirectory

Dim uSleepVbsPath, uPagesCodeFile, uOriginMsgTxtPath, uRecordTxt, uValidResultTxt
uSleepVbsPath = CrtPath & "\sleep.vbs"
uPagesCodeFile = CrtPath & "\tmpFiles\code.txt"
uOriginMsgTxtPath = CrtPath & "\tmpFiles\OriginMsg.txt"
uRecordTxt = CrtPath & "\tmpFiles\record.txt"
uValidResultTxt = CrtPath & "\tmpFiles\validResult.txt"
uInvalidResultTxt = CrtPath & "\tmpFiles\invalidResult.txt"
uAllResultFolder = CrtPath & "\allResults\"

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
    checkStrArray_333 = Array("333","3阶","3速")
    checkStrArray_444 = Array("444","4阶","4速")
    checkStrArray_555 = Array("555","5阶","5速")
    checkStrArray_222 = Array("222","2阶","2速")
    checkStrArray_666 = Array("666","6阶","6速")
    checkStrArray_777 = Array("777","7阶","7速")
    checkStrArray_333bf = Array("3盲","333bf","3bf","3bld")
    checkStrArray_333oh = Array("3单","333oh","3oh","oh","3one")
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

Function initTxtFile(FilePath)
    If Fso.FileExists(FilePath) Then
        Dim TxtFile
        Set TxtFile = Fso.getFile(FilePath)
        TxtFile.Delete
        Set TxtFile = Nothing
    End If    
    Fso.CreateTextFile FilePath, True
End Function

Function getElementValue(elementId)
    getElementValue = document.getElementById(elementId).value
End Function

Sub setElementValue(value)
    document.getElementById(elementId).value = value
End Sub

Sub setInnerHtml(elementId, text)
    document.getElementById(elementId).innerHTML = text
End Sub
