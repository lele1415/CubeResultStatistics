'*************************************************
'****write BR into excel file.
'*************************************************
Function writeBrToExcel(newFilePath, haveNewBr)
    Dim ExcelApp, ExcelBook, ExcelSheet, brLine, aSplitPureResults
    Set ExcelApp = CreateObject("Excel.Application")
    Set ExcelBook= ExcelApp.Workbooks.Open(newFilePath)
    Set ExcelSheet = ExcelBook.Sheets("Sheet2")

    For i = 1 To 57
        For j = 1 To 10
            If i mod 4 = 0 Then
                ExcelSheet.Cells(i, j).Interior.Color = RGB(230,230,230)
            Else
                ExcelSheet.Cells(i, j).Interior.Color = RGB(250,250,250)
            End If
        Next
    Next
            
    If haveNewBr Then
        For i = 0 To brResultCount-1
            aSplitPureResults = Split(brResultData(4,i))
            brLine = getBrLine(brResultData(0,i), brResultData(3,i))

            For j = 1 To 10
                ExcelSheet.Cells(brLine, j).Interior.Color = RGB(255,218,101)
            Next

            ExcelSheet.Cells(brLine, 2).Value = brResultData(2,i)
            ExcelSheet.Cells(brLine, 4).Value = brResultData(1,i)
            ExcelSheet.Cells(brLine, 5).Value = iCompNum & "期"

            If brResultData(3,i) Then
                For j = 0 To UBound(aSplitPureResults)
                    ExcelSheet.Cells(brLine, j + 6).Value = aSplitPureResults(j)
                Next
            End If
        Next
    End If

    ExcelBook.Save
    ExcelBook.Close
    ExcelApp.Quit
    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelApp = Nothing
End Function

        Function getBrLine(sOptName, isAvg)
            Dim tmpLine
            tmpLine = getSeqInAllOptName(sOptName)
            tmpLine = (tmpLine + 1) * 4
            If isAvg Then
                tmpLine = tmpLine + 1
            End If
            getBrLine = tmpLine
        End Function