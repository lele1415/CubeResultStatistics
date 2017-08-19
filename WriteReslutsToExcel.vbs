'*************************************************
'****write final result into excel file.
'*************************************************
Function writeReslutsToExcel()
    Dim oldFilePath, newFilePath
    oldFilePath = uAllResultFolder & (iCompNum - 1) & ".xlsx"
    newFilePath = uAllResultFolder & iCompNum & ".xlsx"
    If Fso.FileExists(oldFilePath) And Not Fso.FileExists(newFilePath) Then
        Fso.CopyFile oldFilePath, newFilePath, False

        Dim ExcelApp, ExcelBook, ExcelSheet
        Set ExcelApp = CreateObject("Excel.Application")
        Set ExcelBook= ExcelApp.Workbooks.Open(newFilePath)

        ExcelApp.DisplayAlerts = False
        ExcelBook.Sheets(1).Delete
        ExcelApp.DisplayAlerts = True

        'Set ExcelSheet = ExcelBook.Sheets("Sheet1")
        Set ExcelSheet = ExcelBook.Sheets.Add

        ExcelSheet.Columns("B").ColumnWidth = 15

        writeAllProjectResult ExcelSheet, aTmpUserText_333, aTmpPureResults_333, aTmpSortedResults_333, "333"
        writeAllProjectResult ExcelSheet, aTmpUserText_444, aTmpPureResults_444, aTmpSortedResults_444, "444"
        writeAllProjectResult ExcelSheet, aTmpUserText_555, aTmpPureResults_555, aTmpSortedResults_555, "555"
        writeAllProjectResult ExcelSheet, aTmpUserText_222, aTmpPureResults_222, aTmpSortedResults_222, "222"
        writeAllProjectResult ExcelSheet, aTmpUserText_333bf, aTmpPureResults_333bf, aTmpSortedResults_333bf, "333bf"
        writeAllProjectResult ExcelSheet, aTmpUserText_333oh, aTmpPureResults_333oh, aTmpSortedResults_333oh, "333oh"
        writeAllProjectResult ExcelSheet, aTmpUserText_333fm, aTmpPureResults_333fm, aTmpSortedResults_333fm, "333fm"
        writeAllProjectResult ExcelSheet, aTmpUserText_py, aTmpPureResults_py, aTmpSortedResults_py, "py"
        writeAllProjectResult ExcelSheet, aTmpUserText_mega, aTmpPureResults_mega, aTmpSortedResults_mega, "mega"
        writeAllProjectResult ExcelSheet, aTmpUserText_sq, aTmpPureResults_sq, aTmpSortedResults_sq, "sq"
        writeAllProjectResult ExcelSheet, aTmpUserText_clock, aTmpPureResults_clock, aTmpSortedResults_clock, "clock"
        writeAllProjectResult ExcelSheet, aTmpUserText_sk, aTmpPureResults_sk, aTmpSortedResults_sk, "sk"
        writeAllProjectResult ExcelSheet, aTmpUserText_666, aTmpPureResults_666, aTmpSortedResults_666, "666"
        writeAllProjectResult ExcelSheet, aTmpUserText_777, aTmpPureResults_777, aTmpSortedResults_777, "777"
        
        ExcelBook.Save
        ExcelBook.Close
        ExcelApp.Quit
        Set ExcelSheet = Nothing
        Set ExcelBook = Nothing
        Set ExcelApp = Nothing

        If brResultCount > 0 Then
            writeBrToExcel newFilePath, True
        Else
            writeBrToExcel newFilePath, False
        End If
    Else
        Msgbox(oldFilePath & Vblf & "not exist" & Vblf & " or" & Vblf & newFilePath & Vblf & "already exist")
    End If

    Msgbox("Write done!")
End Function

        Function writeAllProjectResult(ExcelSheet, aUserText, aPureResults, aSortedResults, sOptName)
            Dim rankNum, isAvgBr, isBestBr
            rankNum = 0
            countExcelRow = countExcelRow + 1
            For i = 1 To 9
                If i = 1 Then
                    ExcelSheet.Cells(countExcelRow, i).HorizontalAlignment = 2
                ElseIf i = 2 Then
                    ExcelSheet.Cells(countExcelRow, i).HorizontalAlignment = 3
                Else
                    ExcelSheet.Cells(countExcelRow, i).HorizontalAlignment = 4
                End If
                ExcelSheet.Cells(countExcelRow, i).Interior.Color = RGB(0,0,0)
                ExcelSheet.Cells(countExcelRow, i).Font.Color = RGB(255,255,255)
            Next

            ExcelSheet.Cells(countExcelRow, 1).Value = getFinalOptName(sOptName)
            ExcelSheet.Cells(countExcelRow, 2).Value = "ID"
            ExcelSheet.Cells(countExcelRow, 3).Value = "最好成绩"
            ExcelSheet.Cells(countExcelRow, 4).Value = "平均成绩"
            ExcelSheet.Cells(countExcelRow, 5).Value = "r1"
            ExcelSheet.Cells(countExcelRow, 6).Value = "r2"
            ExcelSheet.Cells(countExcelRow, 7).Value = "r3"
            ExcelSheet.Cells(countExcelRow, 8).Value = "r4"
            ExcelSheet.Cells(countExcelRow, 9).Value = "r5"

            For i = 0 To safeUBound(aUserText, "writeAllProjectResult", 1)
                Dim aSplitSortedResults, aSplitPureResults
                countExcelRow = countExcelRow + 1
                rankNum = rankNum + 1
                aSplitSortedResults = Split(aSortedResults(i))
                aSplitPureResults = Split(aPureResults(i))
                lenTmpSortedResults = safeUBound(aSplitSortedResults, "writeAllProjectResult", 1)
                isAvgBr = aSplitSortedResults(lenTmpSortedResults - 1)
                isBestBr = aSplitSortedResults(lenTmpSortedResults)

                For j = 1 To 9
                    If j = 2 Then
                        ExcelSheet.Cells(countExcelRow, j).HorizontalAlignment = 2
                    Else
                        ExcelSheet.Cells(countExcelRow, j).HorizontalAlignment = 4
                    End If

                    If j = 4 Then
                        ExcelSheet.Cells(countExcelRow, j).Font.Bold = True
                    End If

                    If j < 3 Then
                        ExcelSheet.Cells(countExcelRow, j).Font.Color = RGB(247,83,9)
                    Else
                        ExcelSheet.Cells(countExcelRow, j).NumberFormat = "@"
                        ExcelSheet.Cells(countExcelRow, j).Font.Color = RGB(0,0,0)
                    End If

                    If isEvenNum(i) Then
                        ExcelSheet.Cells(countExcelRow, j).Interior.Color = RGB(255,255,255)
                    Else
                        ExcelSheet.Cells(countExcelRow, j).Interior.Color = RGB(230,230,230)
                    End If
                Next

                If isBestBr Then
                    ExcelSheet.Cells(countExcelRow, 3).Interior.Color = RGB(255,192,0)
                    saveBrResult sOptName, aUserText(i), aSplitSortedResults(1), False, aPureResults(i)
                End If
                If isAvgBr Then
                    ExcelSheet.Cells(countExcelRow, 4).Interior.Color = RGB(255,192,0)
                    saveBrResult sOptName, aUserText(i), aSplitSortedResults(0), True, aPureResults(i)
                End If

                ExcelSheet.Cells(countExcelRow, 1).Value = rankNum
                ExcelSheet.Cells(countExcelRow, 2).Value = aUserText(i)
                ExcelSheet.Cells(countExcelRow, 3).Value = revertResult(aSplitSortedResults(1), sOptName)
                ExcelSheet.Cells(countExcelRow, 4).Value = revertResult(aSplitSortedResults(0), sOptName)
                For j = 0 To UBound(aSplitPureResults)
                    ExcelSheet.Cells(countExcelRow, j+5).Value = revertResult(aSplitPureResults(j), sOptName)
                Next
            Next

            countExcelRow = countExcelRow + 1
        End Function

                Function saveBrResult(sOptName, sUserText, brResult, isAvg, sPureResult)
                    ReDim Preserve brResultData(4,brResultCount)
                    brResultData(0,brResultCount) = sOptName
                    brResultData(1,brResultCount) = sUserText
                    brResultData(2,brResultCount) = brResult
                    brResultData(3,brResultCount) = isAvg
                    brResultData(4,brResultCount) = sPureResult
                    brResultCount = brResultCount + 1
                End Function