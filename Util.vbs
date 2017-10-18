Const WINDOW_WIDTH = 700
Const WINDOW_HEIGHT = 1000
Sub Window_OnLoad
    Dim ScreenWidth : ScreenWidth = CreateObject("HtmlFile").ParentWindow.Screen.AvailWidth
    Dim ScreenHeight : ScreenHeight = CreateObject("HtmlFile").ParentWindow.Screen.AvailHeight
    Window.MoveTo (ScreenWidth - WINDOW_WIDTH) * 29 / 30 ,(ScreenHeight - WINDOW_HEIGHT) \ 3
    Window.ResizeTo WINDOW_WIDTH, WINDOW_HEIGHT
End Sub

Dim ws, Fso
Set ws=CreateObject("wscript.shell")
Set Fso=CreateObject("Scripting.FileSystemObject")

Dim CrtPath : CrtPath = ws.CurrentDirectory

Dim uSleepVbsPath, uPagesCodeFile, uOriginMsgTxtPath, uRecordTxt, uValidResultTxt
uSleepVbsPath = CrtPath & "\sleep.vbs"
uPagesCodeFile = CrtPath & "\tmpFiles\code.txt"
uOriginMsgTxtPath = CrtPath & "\tmpFiles\OriginMsg.txt"
uRecordTxt = CrtPath & "\tmpFiles\record.txt"
uIgnoreUsersTxt = CrtPath & "\tmpFiles\ignoreUsers.txt"
uValidResultTxt = CrtPath & "\tmpFiles\validResult.txt"
uInvalidResultTxt = CrtPath & "\tmpFiles\invalidResult.txt"
uAllResultFolder = CrtPath & "\allResults\"


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

Sub setElementValue(elementId, value)
    document.getElementById(elementId).value = value
End Sub

Sub setInnerHtml(elementId, text)
    document.getElementById(elementId).innerHTML = text
End Sub

Sub enableElement(elementId)
    document.getElementById(elementId).disabled = ""
End Sub
