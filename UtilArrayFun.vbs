Function msgboxArray(strBegin, strArray)
    Dim tmpStr
    tmpStr = strBegin
    For count = 0 To safeUBound(strArray, "msgboxArray", 1)
        tmpStr = tmpStr & " " & strArray(count)
    Next
    'Msgbox(tmpStr)
End Function

Function getExistInArray(strArray, str)
    Dim countFinded, count
    countFinded = ""
    For count = 0 To safeUBound(strArray, "getExistInArray", 1)
        If str = strArray(count) Then
            countFinded = count
            Exit For
        End If
    Next
    getExistInArray = countFinded
End Function

Function sortIntArray(aArray)
    Dim i, j, lenArray, aTmp
    aTmp = aArray
    lenArray = safeUBound(aTmp, "sortIntArray", 1)
    For i = 0 to lenArray - 1
        For j = i + 1 to lenArray
            If (aTmp(i) - aTmp(j)) > 0 Then
                swapTwoStrings aTmp(i), aTmp(j)
            End If
        Next
    Next
    sortIntArray = aTmp
End Function

FUnction clearArray(strArray)
    For i = 0 To safeUBound(strArray, "clearArray", 1)
        strArray(i) = ""
    Next
End Function

Function getSeqInArrayBySearchStr(sSearchStr, aArray)
    Dim i
    For i = 0 To safeUBound(aArray, "getSeqInArrayBySearchStr", 1)
        If aArray(i) = sSearchStr Then
            getSeqInArrayBySearchStr = i
            Exit For
        End If
    Next
End Function

Function safeUBound(aArray, sFun, dimension)
    If isArray(aArray) Then
        safeUBound = UBound(aArray, dimension)
    Else
        MsgBox("Error: safeUBound ==from==>" & sFun)
    End If
End Function

Function cutArray(aOrigin, iStart, iEnd)
    'MsgBox("iStart="&iStart&" iEnd="&iEnd)
    ReDim aCuted(iEnd - iStart)
    For i = iStart To iEnd
        aCuted(i - iStart) = aOrigin(i)
    Next
    cutArray = aCuted
End Function

Function joinArrayWithSpace(aOrigin)
    Dim sJoined
    sJoined = ""
    For i = 0 To safeUBound(aOrigin, "joinArrayWithSpace", 1)
        sJoined = sJoined & " " & aOrigin(i)
    Next
    joinArrayWithSpace = Trim(sJoined)
End Function

Function coverInArray(aArray, iSeq)
    Dim lenArray
    MsgBox("111")
    lenArray = safeUBound(aArray, "coverInArray", 1)
    If iSeq < lenArray Then
        For i = iSeq To lenArray - 1
            aArray(i) = aArray(i + 1)
        Next
    End If
    MsgBox("222")
    
    ReDim Preserve aArray(lenArray - 1)
    MsgBox("333")
End Function
