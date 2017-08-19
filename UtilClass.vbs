Class PostInfo
    Private mPostNum, mUserID, mMsgStr, mPureResults, mSortedResults, mBestResult, mAvgResult

    Private Sub Class_Initialize
        mPostNum = 0
        mUserID = ""
        mMsgStr = ""
        ReDim Preserve mPureResults(4)
        ReDim Preserve mSortedResults(4)
        mBestResult = ""
        mAvgResult = ""
    End Sub

    Public Property Let PostNum(value)
        mPostNum = value
    End Property

    Public Property Let UserID(value)
        mUserID = value
    End Property

    Public Property Let MsgStr(value)
        mMsgStr = value
    End Property

    Public Property Let PureResults(seq, value)
        If seq >= 0 And seq <= 4 Then
            mPureResults(seq) = value
        Else
            MsgBox("Let PureResults: seq is out of bound. seq = " & seq)
        End If
    End Property

    Public Property Let SortedResults(seq, value)
        If seq >= 0 And seq <= 4 Then
            mSortedResults(seq) = value
        Else
            MsgBox("Let SortedResults: seq is out of bound. seq = " & seq)
        End If
    End Property

    Public Property Get PostNum
        PostNum = mPostNum
    End Property

    Public Property Get UserID
        UserID = mUserID
    End Property

    Public Property Get MsgStr
        MsgStr = mMsgStr
    End Property

    Public Property Get PureResults(seq)
        If seq >= 0 And seq <= 4 Then
            PureResults = mPureResults(seq)
        Else
            MsgBox("Get PureResults: seq is out of bound. seq = " & seq)
        End If
    End Property

    Public Property Get SortedResults(seq)
        If seq >= 0 And seq <= 4 Then
            SortedResults = mSortedResults(seq)
        Else
            MsgBox("Get SortedResults: seq is out of bound. seq = " & seq)
        End If
    End Property
End Class



Class VariableArray
    Private mLength, mArray()

    Private Sub Class_Initialize
        mLength = -1
    End Sub

    Public Property Get Length
        Length = mLength
    End Property

    Public Property Get V(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: Get V(seq) seq is not a number")
            Exit Property
        ELse
            seq = Cint(seq)
        End If

        'MsgBox("seq="&seq&" mLength="&mLength)
        If seq < 0 Or seq > mLength Then
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

        If seq < 0 Or seq > mLength Then
            MsgBox("Error: Let V(seq) seq out of bound")
            Exit Property
        End If

        mArray(seq) = sValue
    End Property

    Public Function Append(value)
        mLength = mLength + 1
        ReDim Preserve mArray(mLength)

        If isObject(value) Then
            Set mArray(mLength) = value
        ELse
            mArray(mLength) = value
        End If
    End Function

    Public Function ResetArray()
        mLength = -1
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
            mLength = mLength + 1
            ReDim Preserve mArray(mLength)
            mArray(mLength) = newArray(i)
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

        If (seq1 < 0 Or seq1 > mLength) Or (seq2 < 0 Or seq2 > mLength) Then
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

        If seq < 0 Or seq > mLength Then
            MsgBox("Error: PopBySeq(seq) seq out of bound")
            Exit Function
        End If

        If seq <> mLength Then
            Dim i
            For i = seq To mLength - 1
                mArray(i) = mArray(i + 1)
            Next
        End If

        mLength = mLength - 1
        ReDim Preserve mArray(mLength)
    End Function

    Public Function MoveToTop(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: MoveToTop(seq) seq is not a number")
            Exit Function
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mLength Then
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

        If seq < 0 Or seq > mLength Then
            MsgBox("Error: MoveToEnd(seq) seq out of bound")
            Exit Function
        End If

        If seq = 0 Then Exit Function

        Dim i, sValueToBeMove
        If isObject(mArray(seq)) Then
            Set sValueToBeMove = mArray(seq)
            For i = seq To mLength - 1
                Set mArray(i) = mArray(i + 1)
            Next
            Set mArray(mLength) = sValueToBeMove
        Else
            sValueToBeMove = mArray(seq)
            For i = seq To mLength - 1
                mArray(i) = mArray(i + 1)
            Next
            mArray(mLength) = sValueToBeMove
        End If
    End Function

    Public Function IsExistInArray(value)
        If mLength = -1 Then
            IsExistInArray = False
            Exit Function
        End If

        Dim i
        For i = 0 To mLength
            If StrComp(mArray(i), value) = 0 Then
                IsExistInArray = True
                Exit Function
            End If
        Next
        IsExistInArray = False
    End Function

    Public Function IsExistInObject(value, seq)
        If mLength = -1 Then
            IsExistInObject = False
            Exit Function
        End If

        Dim i
        For i = 0 To mLength
            If StrComp(mArray(i).V(seq), value) = 0 Then
                IsExistInObject = True
                Exit Function
            End If
        Next
        IsExistInObject = False
    End Function

    Public Function SortArray()
        If mLength = -1 Then
            'MsgBox("Error: SortArray() mLength <= 0, no need to sort")
            Exit Function
        End If

        Dim i, j
        For i = 0 To mLength - 1
            For j = i + 1 To mLength
                If StrComp(mArray(i), mArray(j)) > 0 Then
                    Dim sTmp : sTmp = mArray(i) : mArray(i) = mArray(j) : mArray(j) = sTmp
                End If
            Next
        Next
    End Function

    Public Function ToString()
        If mLength <> -1 Then
            Dim i, sTmp
            sTmp = "v(0) = " & mArray(0)
            If mLength > 0 Then
                For i = 1 To mLength
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