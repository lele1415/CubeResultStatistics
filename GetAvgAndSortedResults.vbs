'*************************************************
'****get avg and sorted results.
'*************************************************
Function getAvgAndSortedResults(sOptName, sPureResult, isValid)
    If isValid Then
        Dim aTmp, aPureResults, lenResults, iAvgResult, iEnd
        aPureResults = Split(sPureResult)
        aTmp = sortIntArray(aPureResults)
        lenResults = safeUBound(aTmp, "getAvgAndSortedResults", 1)

        If lenResults = 4 Then
            iEnd = 3
        Else
            iEnd = 2
        End If

        If aTmp(iEnd) <> 9999.99 Then
            iAvgResult = Round(((0 + aTmp(iEnd-2) + aTmp(iEnd-1) + aTmp(iEnd)) / 3), 2)
            iAvgResult = FormatNumber(iAvgResult, 2, , , 0)
        Else
            iAvgResult = 9999.99
        End If

        Dim sBrFlag
        sBrFlag = compareToRecord(sOptName, iAvgResult, aTmp(0))

        getAvgAndSortedResults = iAvgResult & " " & joinArrayWithSpace(aTmp) & " " & sBrFlag
    Else
        getAvgAndSortedResults = ""
    End If
End Function

Function loadAllRecords()
    Dim oTxt, count
    Set oTxt = Fso.OpenTextFile(uRecordTxt, 1)
    count = 0

    Do Until oTxt.AtEndOfStream
        aAllRecords(count) = oTxt.ReadLine
        count = count + 1
    Loop

    oTxt.Close
    Set oTxt = Nothing
End Function

Function compareToRecord(sOptName, iAvgResult, iBestResult)
    Dim iOptCount, aTmpRecord, iAvgRecord, iBestRecord, isAvgBr, isBestBr
    iOptCount = getSeqInAllOptName(sOptName)
    aTmpRecord = Split(aAllRecords(iOptCount))
    iAvgRecord = aTmpRecord(2)
    iBestRecord = aTmpRecord(1)
    isAvgBr = False
    isBestBr = False

    If sOptName <> "333fm" Then
        iAvgRecord = formatResultStr(iAvgRecord)
        iBestRecord = formatResultStr(iBestRecord)
    End If

    If iAvgResult - iAvgRecord <= 0 Then
        isAvgBr = True
    End If

    If iBestResult - iBestRecord <= 0 Then
        isBestBr = True
    End If

    compareToRecord = isAvgBr & " " & isBestBr
End Function
