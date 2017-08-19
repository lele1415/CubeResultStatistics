Class PostInfo
    Private mPostNum, mPostUser, mPostMsg

    Private Sub Class_Initialize
        mPostNum = 0
        mPostUser = ""
        mPostMsg = ""
    End Sub

    Public Property Let PostNum(value)
        mPostNum = value
    End Property

    Public Property Let PostUser(value)
        mPostUser = value
    End Property

    Public Property Let PostMsg(value)
        mPostMsg = value
    End Property

    Public Property Get PostNum
        PostNum = mPostNum
    End Property

    Public Property Get PostUser
        PostUser = mPostUser
    End Property

    Public Property Get PostMsg
        PostMsg = mPostMsg
    End Property
End Class



Class ResultInfo
    Private mResultOwner, mResultText, mResultOpt, mPureResults, mSortedResults, mBestResult, mAvgResult

    Private Sub Class_Initialize
        mResultOwner = ""
        mResultText = ""
        ReDim Preserve mPureResults(4)
        ReDim Preserve mSortedResults(4)
        mBestResult = ""
        mAvgResult = ""
    End Sub

    Public Property Let ResultOwner(value)
        mResultOwner = value
    End Property

    Public Property Let ResultText(value)
        mResultText = value
    End Property

    Public Property Let ResultOpt(value)
        mResultOpt = value
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

    Public Property Let BestResult(value)
        mBestResult = value
    End Property

    Public Property Let AvgResult(value)
        mAvgResult = value
    End Property

    Public Property Get ResultOwner
        ResultOwner = mResultOwner
    End Property

    Public Property Get ResultText
        ResultText = mResultText
    End Property

    Public Property Get ResultOpt
        ResultOpt = mResultOpt
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

    Public Property Get BestResult
        BestResult = mBestResult
    End Property

    Public Property Get AvgResult
        AvgResult = mAvgResult
    End Property
End Class



Class VariableArray
    Private mBound, mArray()

    Private Sub Class_Initialize
        mBound = -1
    End Sub

    Public Property Get Bound
        Bound = mBound
    End Property

    Public Property Get V(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: Get V(seq) seq is not a number")
            Exit Property
        ELse
            seq = Cint(seq)
        End If

        'MsgBox("seq="&seq&" mBound="&mBound)
        If seq < 0 Or seq > mBound Then
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

        If seq < 0 Or seq > mBound Then
            MsgBox("Error: Let V(seq) seq out of bound")
            Exit Property
        End If

        mArray(seq) = sValue
    End Property

    Public Sub Append(value)
        mBound = mBound + 1
        ReDim Preserve mArray(mBound)

        If isObject(value) Then
            Set mArray(mBound) = value
        ELse
            mArray(mBound) = value
        End If
    End Sub

    Public Sub ResetArray()
        mBound = -1
    End Sub

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
            mBound = mBound + 1
            ReDim Preserve mArray(mBound)
            mArray(mBound) = newArray(i)
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

        If (seq1 < 0 Or seq1 > mBound) Or (seq2 < 0 Or seq2 > mBound) Then
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

    Public Sub PopBySeq(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: PopBySeq(seq) seq is not a number")
            Exit Sub
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mBound Then
            MsgBox("Error: PopBySeq(seq) seq out of bound")
            Exit Sub
        End If

        If seq <> mBound Then
            Dim i
            For i = seq To mBound - 1
                mArray(i) = mArray(i + 1)
            Next
        End If

        mBound = mBound - 1
        ReDim Preserve mArray(mBound)
    End Sub

    Public Sub MoveToTop(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: MoveToTop(seq) seq is not a number")
            Exit Sub
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mBound Then
            MsgBox("Error: MoveToTop(seq) seq out of bound")
            Exit Sub
        End If

        If seq = 0 Then Exit Sub

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
    End Sub

    Public Sub MoveToEnd(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: MoveToEnd(seq) seq is not a number")
            Exit Sub
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mBound Then
            MsgBox("Error: MoveToEnd(seq) seq out of bound")
            Exit Sub
        End If

        If seq = 0 Then Exit Sub

        Dim i, sValueToBeMove
        If isObject(mArray(seq)) Then
            Set sValueToBeMove = mArray(seq)
            For i = seq To mBound - 1
                Set mArray(i) = mArray(i + 1)
            Next
            Set mArray(mBound) = sValueToBeMove
        Else
            sValueToBeMove = mArray(seq)
            For i = seq To mBound - 1
                mArray(i) = mArray(i + 1)
            Next
            mArray(mBound) = sValueToBeMove
        End If
    End Sub

    Public Function IsExistInArray(value)
        If mBound = -1 Then
            IsExistInArray = False
            Exit Function
        End If

        Dim i
        For i = 0 To mBound
            If StrComp(mArray(i), value) = 0 Then
                IsExistInArray = True
                Exit Function
            End If
        Next
        IsExistInArray = False
    End Function

    Public Function IsExistInObject(value, seq)
        If mBound = -1 Then
            IsExistInObject = False
            Exit Function
        End If

        Dim i
        For i = 0 To mBound
            If StrComp(mArray(i).V(seq), value) = 0 Then
                IsExistInObject = True
                Exit Function
            End If
        Next
        IsExistInObject = False
    End Function

    Public Sub SortArray()
        If mBound = -1 Then
            'MsgBox("Error: SortArray() mBound <= 0, no need to sort")
            Exit Sub
        End If

        Dim i, j
        For i = 0 To mBound - 1
            For j = i + 1 To mBound
                If StrComp(mArray(i), mArray(j)) > 0 Then
                    Dim sTmp : sTmp = mArray(i) : mArray(i) = mArray(j) : mArray(j) = sTmp
                End If
            Next
        Next
    End Sub

    Public Function ToString()
        If mBound <> -1 Then
            Dim i, sTmp
            sTmp = "v(0) = " & mArray(0)
            If mBound > 0 Then
                For i = 1 To mBound
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