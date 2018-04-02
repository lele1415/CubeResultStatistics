'*************************************************
'****write final result into excel file.
'*************************************************
Const ID_SAVE_RESULTS_STATUS = "save_results_status"

Dim iExcelRow
Dim vaAllBrResults : Set vaAllBrResults = New VariableArray
Dim uPostInfoTxtPath : uPostInfoTxtPath = CrtPath & "\tmpFiles\ResultInfo.txt"

Sub onClickSaveAllResult()
    Call setInnerHtml(ID_SAVE_RESULTS_STATUS, "正在保存...")
    idTimer = window.setTimeout("saveAllResult()", 0, "VBScript")
End Sub

Sub saveAllResult()
    window.clearTimeout(idTimer)

    saveValidResultInfoToTxt()
    writeReslutsToExcel()

    Call setInnerHtml(ID_SAVE_RESULTS_STATUS, "已完成")
End Sub

        Sub saveValidResultInfoToTxt()
            initTxtFile(uPostInfoTxtPath)
            Dim oTxt, i, obj
            Set oTxt = Fso.OpenTextFile(uPostInfoTxtPath, 8, False, True)

            For i = 0 To vaAllValidResultInfo.Bound
                Set obj = vaAllValidResultInfo.V(i)
                oTxt.WriteLine(obj.PostNum)
                oTxt.WriteLine(obj.ResultOwner)
                oTxt.WriteLine(vaOptInfo.V(obj.ResultOptSeq).FullName)
                oTxt.WriteLine(obj.ResultText)
                oTxt.WriteLine(obj.PureResults)
                oTxt.WriteLine(obj.SortedResults)
                oTxt.WriteLine(obj.BestResult)
                oTxt.WriteLine(obj.AvgResult)
                oTxt.WriteLine(obj.IsBestBr)
                oTxt.WriteLine(obj.IsAvgBr)
                oTxt.WriteLine()
            Next

            oTxt.Close
            Set oTxt = Nothing
        End Sub

Sub writeReslutsToExcel()
    Dim oldFilePath, newFilePath
    If iCompNum = 0 Then iCompNum = InputBox("input iCompNum")

    oldFilePath = uAllResultFolder & (iCompNum - 1) & ".xlsx"
    newFilePath = uAllResultFolder & iCompNum & ".xlsx"
    If Not Fso.FileExists(oldFilePath) Or Fso.FileExists(newFilePath) Then
        Msgbox(oldFilePath & Vblf & "not exist" & Vblf & " or" & Vblf & newFilePath & Vblf & "already exist")
        Exit Sub
    End If

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

    Call writeAllProjectResult(ExcelSheet, vaAllResults_333, OPT_SEQ_333)
    Call writeAllProjectResult(ExcelSheet, vaAllResults_222, OPT_SEQ_222)
    Call writeAllProjectResult(ExcelSheet, vaAllResults_444, OPT_SEQ_444)
    Call writeAllProjectResult(ExcelSheet, vaAllResults_555, OPT_SEQ_555)
    Call writeAllProjectResult(ExcelSheet, vaAllResults_666, OPT_SEQ_666)
    Call writeAllProjectResult(ExcelSheet, vaAllResults_777, OPT_SEQ_777)
    Call writeAllProjectResult(ExcelSheet, vaAllResults_3bf, OPT_SEQ_3bf)
    Call writeAllProjectResult(ExcelSheet, vaAllResults_3fm, OPT_SEQ_3fm)
    Call writeAllProjectResult(ExcelSheet, vaAllResults_3oh, OPT_SEQ_3oh)
    Call writeAllProjectResult(ExcelSheet, vaAllResults_3wf, OPT_SEQ_3wf)
    Call writeAllProjectResult(ExcelSheet, vaAllResults_mega, OPT_SEQ_mega)
    Call writeAllProjectResult(ExcelSheet, vaAllResults_py, OPT_SEQ_py)
    Call writeAllProjectResult(ExcelSheet, vaAllResults_clk, OPT_SEQ_clk)
    Call writeAllProjectResult(ExcelSheet, vaAllResults_sk, OPT_SEQ_sk)
    Call writeAllProjectResult(ExcelSheet, vaAllResults_sq, OPT_SEQ_sq)
    Call writeAllProjectResult(ExcelSheet, vaAllResults_4bf, OPT_SEQ_4bf)
    Call writeAllProjectResult(ExcelSheet, vaAllResults_5bf, OPT_SEQ_5bf)
    Call writeAllProjectResult(ExcelSheet, vaAllResults_3mb, OPT_SEQ_3mb)
    
    If vaAllBrResults.Bound <> -1 Then
        Call writeBrToExcel(ExcelApp, ExcelBook, ExcelSheet, True)
        Call saveAllBrResultToTxt()
    Else
        Call writeBrToExcel(ExcelApp, ExcelBook, ExcelSheet, False)
    End If

    ExcelBook.Save
    ExcelBook.Close
    ExcelApp.Quit
    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelApp = Nothing
End Sub

        Sub writeAllProjectResult(ExcelSheet, vaObj, iOptSeq)
            Dim iRankNum, isAvgBr, isBestBr
            iRankNum = 0
            iExcelRow = iExcelRow + 1

            Call writeTitleForOpt(ExcelSheet, iOptSeq)
            

            Dim i, k, oResult, aTmp
            For i = 0 To vaObj.Bound
                Set oResult = vaObj.V(i)
                iExcelRow = iExcelRow + 1
                iRankNum = iRankNum + 1

                Call setStyleForResult(ExcelSheet, i, oResult, iOptSeq)
                Call checkIsBr(ExcelSheet, oResult)

                ExcelSheet.Cells(iExcelRow, 1).Value = iRankNum
                ExcelSheet.Cells(iExcelRow, 2).Value = oResult.ResultOwner

                If iOptSeq <> OPT_SEQ_3mb Then
                    ExcelSheet.Cells(iExcelRow, 3).Value = revertResult(oResult.BestResult, iOptSeq)
                Else
                    Dim aPureResults
                    aPureResults = Split(oResult.PureResults, " ")
                    ExcelSheet.Cells(iExcelRow, 3).Value = aPureResults(0) & "/" & aPureResults(1)
                End If

                If iOptSeq < 15 Then
                    ExcelSheet.Cells(iExcelRow, 4).Value = revertResult(oResult.AvgResult, iOptSeq)
                ElseIf iOptSeq = OPT_SEQ_3mb Then
                    ExcelSheet.Cells(iExcelRow, 4).Value = revertResult(oResult.AvgResult, iOptSeq)
                Else
                    ExcelSheet.Cells(iExcelRow, 4).Value = ""
                End If

                If iOptSeq <> OPT_SEQ_3mb Then
                    aTmp = Split(oResult.PureResults)
                    For k = 0 To UBound(aTmp)
                        ExcelSheet.Cells(iExcelRow, k + 5).Value = revertResult(aTmp(k), iOptSeq)
                    Next
                End If
            Next

            iExcelRow = iExcelRow + 1
        End Sub

                Sub writeTitleForOpt(ExcelSheet, iOptSeq)
                    Dim oOptInfo
                    Set oOptInfo = vaOptInfo.V(iOptSeq)

                    For i = 1 To 9
                        If i = 1 Then
                            ExcelSheet.Cells(iExcelRow, i).HorizontalAlignment = 2
                        ElseIf i = 2 Then
                            ExcelSheet.Cells(iExcelRow, i).HorizontalAlignment = 3
                        Else
                            ExcelSheet.Cells(iExcelRow, i).HorizontalAlignment = 4
                        End If
                        ExcelSheet.Cells(iExcelRow, i).Interior.Color = RGB(0,0,0)
                        ExcelSheet.Cells(iExcelRow, i).Font.Color = RGB(255,255,255)
                    Next

                    ExcelSheet.Cells(iExcelRow, 1).Value = oOptInfo.FullName
                    ExcelSheet.Cells(iExcelRow, 2).Value = "ID"
                    If iOptSeq <> OPT_SEQ_3mb Then
                        ExcelSheet.Cells(iExcelRow, 3).Value = "最好成绩"
                        If iOptSeq <> OPT_SEQ_4bf And _
                                iOptSeq <> OPT_SEQ_5bf Then
                            ExcelSheet.Cells(iExcelRow, 4).Value = "平均成绩"
                        End If
                        ExcelSheet.Cells(iExcelRow, 5).Value = "r1"
                        ExcelSheet.Cells(iExcelRow, 6).Value = "r2"
                        ExcelSheet.Cells(iExcelRow, 7).Value = "r3"
                        If oOptInfo.NeedNum = 5 Then
                            ExcelSheet.Cells(iExcelRow, 8).Value = "r4"
                            ExcelSheet.Cells(iExcelRow, 9).Value = "r5"
                        End If
                    Else
                        ExcelSheet.Cells(iExcelRow, 3).Value = "成功/总数"
                        ExcelSheet.Cells(iExcelRow, 4).Value = "用时"
                    End If
                End Sub

                Sub setStyleForResult(ExcelSheet, iCrtSeq, oResult, iOptSeq)
                    Dim iColumnBold
                    If iOptSeq = OPT_SEQ_3bf Or _
                            iOptSeq = OPT_SEQ_4bf Or _
                            iOptSeq = OPT_SEQ_5bf Or _
                            iOptSeq = OPT_SEQ_3mb Then
                        iColumnBold = 3
                    Else
                        iColumnBold = 4
                    End If

                    Dim col
                    For col = 1 To 9
                        If col = 2 Then
                            ExcelSheet.Cells(iExcelRow, col).HorizontalAlignment = 2
                        Else
                            ExcelSheet.Cells(iExcelRow, col).HorizontalAlignment = 4
                        End If

                        If col = iColumnBold Then
                            ExcelSheet.Cells(iExcelRow, col).Font.Bold = True
                        End If

                        If col < 3 Then
                            ExcelSheet.Cells(iExcelRow, col).Font.Color = RGB(247,83,9)
                        Else
                            ExcelSheet.Cells(iExcelRow, col).NumberFormat = "@"
                            ExcelSheet.Cells(iExcelRow, col).Font.Color = RGB(0,0,0)
                        End If

                        If isEvenNum(iCrtSeq) Then
                            ExcelSheet.Cells(iExcelRow, col).Interior.Color = RGB(255,255,255)
                        Else
                            ExcelSheet.Cells(iExcelRow, col).Interior.Color = RGB(230,230,230)
                        End If
                    Next
                End Sub

                Sub checkIsBr(ExcelSheet, oResult)
                    If oResult.IsBestBr Or oResult.IsAvgBr Then
                        vaAllBrResults.Append(oResult)

                        If oResult.IsBestBr Then
                            ExcelSheet.Cells(iExcelRow, 3).Interior.Color = RGB(255,192,0)
                        End If

                        If oResult.IsAvgBr Then
                            ExcelSheet.Cells(iExcelRow, 4).Interior.Color = RGB(255,192,0)
                        End If
                    End If
                End Sub

Sub writeBrToExcel(ExcelApp, ExcelBook, ExcelSheet, haveNewBr)
    Dim aSplitPureResults
    Set ExcelSheet = ExcelBook.Sheets("Sheet2")

    
    Call clearColorForBrSheet(ExcelSheet)
            
    If haveNewBr Then
        Dim i, oResult, iBaseLine
        For i = 0 To vaAllBrResults.Bound
            Set oResult = vaAllBrResults.V(i)
            If oResult.ResultOptSeq < 16 Then
                iBaseLine = (oResult.ResultOptSeq + 1) * 4
            Else
                iBaseLine = 64 + (oResult.ResultOptSeq + 1 - 16) * 3
            End If

            If oResult.IsBestBr Then
                If oResult.ResultOptSeq <> OPT_SEQ_3mb Then
                    Call writeBrLine(ExcelSheet, iBaseLine, oResult.ResultOptSeq, oResult.BestResult, oResult.ResultOwner, "")
                Else
                    Dim aPureResults, tmpBestResult
                    aPureResults = Split(oResult.PureResults, " ")
                    tmpBestResult = aPureResults(0) & "/" & aPureResults(1)
                    Call writeBrLine(ExcelSheet, iBaseLine, oResult.ResultOptSeq, tmpBestResult, oResult.ResultOwner, oResult.AvgResult)
                End If
            End If

            If oResult.IsAvgBr Then
                Call writeBrLine(ExcelSheet, iBaseLine + 1, oResult.ResultOptSeq, oResult.AvgResult, oResult.ResultOwner, oResult.PureResults)
            End If
        Next
    End If
End Sub

        Sub clearColorForBrSheet(ExcelSheet)
            Dim i, j
            For i = 1 To 70
                For j = 1 To 10
                    If i mod 4 = 0 Then
                        ExcelSheet.Cells(i, j).Interior.Color = RGB(230,230,230)
                    Else
                        ExcelSheet.Cells(i, j).Interior.Color = RGB(250,250,250)
                    End If
                Next
            Next
        End Sub

        Sub writeBrLine(ExcelSheet, iLine, iOptSeq, iResult, sResultOwner, sPureResult)
            Dim j
            For j = 1 To 10
                ExcelSheet.Cells(iLine, j).Interior.Color = RGB(255,218,101)
            Next

            Dim brResult
            If iOptSeq <> OPT_SEQ_3mb Then
                brResult = revertResult(iResult, iOptSeq)
            Else
                brResult = iResult
            End If
            ExcelSheet.Cells(iLine, 2).Value = brResult
            ExcelSheet.Cells(iLine, 4).Value = sResultOwner
            ExcelSheet.Cells(iLine, 5).Value = iCompNum & "期"

            If sPureResult <> "" Then
                Dim aTmp
                aTmp = Split(sPureResult)
                For j = 0 To UBound(aTmp)
                    ExcelSheet.Cells(iLine, j + 6).Value = revertResult(aTmp(j), iOptSeq)
                Next
            End If
        End Sub

Sub saveAllBrResultToTxt()

    Dim i, oBrResult, iOptSeq
    For i = 0 To vaAllBrResults.Bound
        Set oBrResult = vaAllBrResults.V(i)
        iOptSeq = oBrResult.ResultOptSeq

        If oBrResult.IsBestBr Then
            vaOptInfo.V(iOptSeq).BestRecord = oBrResult.BestResult
        End If

        If oBrResult.IsAvgBr Then
            vaOptInfo.V(iOptSeq).AvgRecord = oBrResult.AvgResult
        End If
    Next

    
    initTxtFile(uRecordTxt)
    Dim oTxt, oOptInfo
    Set oTxt = Fso.OpenTextFile(uRecordTxt, 8, False, True)

    For i = 0 To vaOptInfo.Bound
        Set oOptInfo = vaOptInfo.V(i)
        oTxt.WriteLine(oOptInfo.FullName & " " & revertResult(oOptInfo.BestRecord, i) & " " & revertResult(oOptInfo.AvgRecord, i))
    Next

    oTxt.Close
    Set oTxt = Nothing
End Sub
