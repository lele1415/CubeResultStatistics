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

Class VariableArray
    Private mUBd, mArray()

    Private Sub Class_Initialize
        mUBd = -1
    End Sub

    Public Property Get UBd
        UBd = mUBd
    End Property

    Public Property Get V(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: Get V(seq) seq is not a number")
            Exit Property
        ELse
            seq = Cint(seq)
        End If

        'MsgBox("seq="&seq&" mUBd="&mUBd)
        If seq < 0 Or seq > mUBd Then
            MsgBox("Error: Get V(seq) seq out of bound")
            Exit Property
        End If

        If isObject(mArray(seq)) Then
            Set V = mArray(seq)
        Else
            V = mArray(seq)
        End If
    End Property

    Public Property Let V(seq, sValue)
        If Not isNumeric(seq) Then
            MsgBox("Error: Let V(seq) seq is not a number")
            Exit Property
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mUBd Then
            MsgBox("Error: Let V(seq) seq out of bound")
            Exit Property
        End If

        mArray(seq) = sValue
    End Property

    Public Function Append(value)
        mUBd = mUBd + 1
        ReDim Preserve mArray(mUBd)

        If isObject(value) Then
            Set mArray(mUBd) = value
        ELse
            mArray(mUBd) = value
        End If
    End Function

    Public Function ResetArray()
        mUBd = -1
    End Function

    Public Property Get InnerArray
        InnerArray = mArray
    End Property

    Public Property Let InnerArray(newArray)
        If Not isArray(newArray) Then
            MsgBox("Error: Set InnerArray(newArray) newArray is not array")
            Exit Property
        End If

        Dim i
        For i = 0 To UBound(newArray)
            mUBd = mUBd + 1
            ReDim Preserve mArray(mUBd)
            mArray(mUBd) = newArray(i)
        Next
    End Property

    Public Sub SwapTwoValues(seq1, seq2)
        If Not isNumeric(seq1) Or Not isNumeric(seq2) Then
            MsgBox("Error: SwapTwoValues(seq1, seq2) seq1 or seq2 is not a number")
            Exit Sub
        ELse
            seq1 = Cint(seq1)
            seq2 = Cint(seq2)
        End If

        If (seq1 < 0 Or seq1 > mUBd) Or (seq2 < 0 Or seq2 > mUBd) Then
            MsgBox("Error: SwapTwoValues(seq1, seq2) seq1 or seq2 out of bound")
            Exit Sub
        End If

        If isObject(mArray(seq1)) And isObject(mArray(seq2)) Then
            Dim oTmp1, oTmp2
            Set oTmp1 = mArray(seq1)
            Set oTmp2 = mArray(seq2)

            Set mArray(seq1) = Nothing
            Set mArray(seq2) = Nothing

            Set mArray(seq1) = oTmp2
            Set mArray(seq2) = oTmp1
        Else
            Dim sTmp
            sTmp = mArray(seq1)
            mArray(seq1) = mArray(seq2)
            mArray(seq2) = sTmp
        End If
    End Sub

    Public Function PopBySeq(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: PopBySeq(seq) seq is not a number")
            Exit Function
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mUBd Then
            MsgBox("Error: PopBySeq(seq) seq out of bound")
            Exit Function
        End If

        If seq <> mUBd Then
            Dim i
            For i = seq To mUBd - 1
                mArray(i) = mArray(i + 1)
            Next
        End If

        mUBd = mUBd - 1
        ReDim Preserve mArray(mUBd)
    End Function

    Public Function MoveToTop(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: MoveToTop(seq) seq is not a number")
            Exit Function
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mUBd Then
            MsgBox("Error: MoveToTop(seq) seq out of bound")
            Exit Function
        End If

        If seq = 0 Then Exit Function

        Dim i, sValueToBeMove
        If isObject(mArray(seq)) Then
            Set sValueToBeMove = mArray(seq)
            For i = seq To 1 Step -1
                Set mArray(i) = mArray(i - 1)
                Set mArray(0) = sValueToBeMove
            Next
        Else
            sValueToBeMove = mArray(seq)
            For i = seq To 1 Step -1
                mArray(i) = mArray(i - 1)
                mArray(0) = sValueToBeMove
            Next
        End If
    End Function

    Public Function MoveToEnd(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: MoveToEnd(seq) seq is not a number")
            Exit Function
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mUBd Then
            MsgBox("Error: MoveToEnd(seq) seq out of bound")
            Exit Function
        End If

        If seq = 0 Then Exit Function

        Dim i, sValueToBeMove
        If isObject(mArray(seq)) Then
            Set sValueToBeMove = mArray(seq)
            For i = seq To mUBd - 1
                Set mArray(i) = mArray(i + 1)
            Next
            Set mArray(mUBd) = sValueToBeMove
        Else
            sValueToBeMove = mArray(seq)
            For i = seq To mUBd - 1
                mArray(i) = mArray(i + 1)
            Next
            mArray(mUBd) = sValueToBeMove
        End If
    End Function

    Public Function IsExistInArray(value)
        If mUBd = -1 Then
            IsExistInArray = False
            Exit Function
        End If

        Dim i
        For i = 0 To mUBd
            If StrComp(mArray(i), value) = 0 Then
                IsExistInArray = True
                Exit Function
            End If
        Next
        IsExistInArray = False
    End Function

    Public Function IsExistInObject(value, seq)
        If mUBd = -1 Then
            IsExistInObject = False
            Exit Function
        End If

        Dim i
        For i = 0 To mUBd
            If StrComp(mArray(i).V(seq), value) = 0 Then
                IsExistInObject = True
                Exit Function
            End If
        Next
        IsExistInObject = False
    End Function

    Public Function SortArray()
        If mUBd = -1 Then
            MsgBox("Error: SortArray() mUBd <= 0, no need to sort")
            Exit Function
        End If

        Dim i, j
        For i = 0 To mUBd - 1
            For j = i + 1 To mUBd
                If StrComp(mArray(i), mArray(j)) > 0 Then
                    Dim sTmp : sTmp = mArray(i) : mArray(i) = mArray(j) : mArray(j) = sTmp
                End If
            Next
        Next
    End Function

    Public Function ToString()
        If mUBd <> -1 Then
            Dim i, sTmp
            sTmp = "v(0) = " & mArray(0)
            If mUBd > 0 Then
                For i = 1 To mUBd
                    If isArray(mArray(i)) Then
                        sTmp = sTmp & Vblf & "v(" & i & ") = " & join(mArray(i))
                    ElseIf isObject(mArray(i)) Then
                        sTmp = sTmp & Vblf & "v(" & i & ") = [Object]"
                    Else
                        sTmp = sTmp & Vblf & "v(" & i & ") = " & mArray(i)
                    End If
                Next
            End If
            ToString = sTmp
        Else
            MsgBox("Error: ToString() mArray has no element")
        End If
    End Function
End Class