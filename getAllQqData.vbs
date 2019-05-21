'*************************************************
'****get data array.
'*************************************************
Const ID_QQ_DATA_FOLDER = "qq_data_folder"

Dim sQqFilePath
Dim iFirstQqDataIndex

Sub onClickGetAllQqData()
    idTimer = window.setTimeout("getAllQqData()", 10, "VBScript")
End Sub

Sub getAllQqData()
    window.clearTimeout(idTimer)

    sQqFilePath = getElementValue(ID_QQ_DATA_FOLDER) + "\qq.xlsx"
    If Not Fso.FileExists(sQqFilePath) Then
        MsgBox("file not exists!")
        Exit Sub
    End If

    iFirstQqDataIndex = vaAllPostInfo.Bound + 1

    Dim ExcelApp, ExcelBook, ExcelSheet
    Set ExcelApp = CreateObject("Excel.Application")
    Set ExcelBook= ExcelApp.Workbooks.Open(sQqFilePath)
    Set ExcelSheet = ExcelBook.Sheets("Sheet1")

    Dim row, name, opt, results, oTmp
    row = 2
    
    Do While Trim(ExcelSheet.Cells(row, 2).Value) <> ""
        name = ExcelSheet.Cells(row, 2).Value
        opt = ExcelSheet.Cells(row, 4).Value
        results = ExcelSheet.Cells(row, 7).Value

        Set oTmp = New PostInfo
        oTmp.PostNum = 9999 + row
        oTmp.PostUser = name
        oTmp.PostMsg = opt & " " & results

        Dim oNew : Set oNew = oTmp
        vaAllPostInfo.Append(oNew)

        row = row + 1
    Loop

    ExcelBook.Close
    ExcelApp.Quit
    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelApp = Nothing

    Call saveQqDataTxt()

    MsgBox("getAllQqData done!")
End Sub

Sub saveQqDataTxt()
    if vaAllPostInfo.Bound < iFirstQqDataIndex Then Exit Sub

    initTxtFile(uOriginQqTxtPath)
    Dim oTxt, i
    Set oTxt = Fso.OpenTextFile(uOriginQqTxtPath, 8, False, True)

    For i = iFirstQqDataIndex To vaAllPostInfo.Bound
        oTxt.WriteLine(vaAllPostInfo.V(i).PostNum)
        oTxt.WriteLine(vaAllPostInfo.V(i).PostUser)
        oTxt.WriteLine(vaAllPostInfo.V(i).PostMsg)
        oTxt.WriteLine()
    Next

    oTxt.Close
    Set oTxt = Nothing
End Sub
