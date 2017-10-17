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



Class OptionInfo
    Private mFullName, mSeq, mNeedNum, mBestRecord, mAvgRecord, mCount

    Private Sub Class_Initialize
        mFullName = ""
        mSeq = ""
        mNeedNum = ""
        mBestRecord = ""
        mAvgRecord = ""
        mCount = -1
    End Sub

    Public Property Let FullName(value)
        mFullName = value
    End Property

    Public Property Let Seq(value)
        mSeq = value
    End Property

    Public Property Let NeedNum(value)
        mNeedNum = value
    End Property

    Public Property Let BestRecord(value)
        mBestRecord = value
    End Property

    Public Property Let AvgRecord(value)
        mAvgRecord = value
    End Property

    Public Sub CountPlus()
        mCount = mCount + 1
    End Sub

    Public Property Get FullName
        FullName = mFullName
    End Property

    Public Property Get Seq
        Seq = mSeq
    End Property

    Public Property Get NeedNum
        NeedNum = mNeedNum
    End Property

    Public Property Get BestRecord
        BestRecord = mBestRecord
    End Property

    Public Property Get AvgRecord
        AvgRecord = mAvgRecord
    End Property

    Public Property Get Count
        Count = mCount
    End Property
End Class



Class OptLocationInfo
    Private mOptSeq, mOptLocation

    Public Property Let OptSeq(value)
        mOptSeq = value
    End Property

    Public Property Get OptSeq()
        OptSeq = mOptSeq
    End Property

    Public Property Let OptLocation(value)
        mOptLocation = value
    End Property

    Public Property Get OptLocation()
        OptLocation = mOptLocation
    End Property
End Class



Class ValidResultInfo
    Private mPostNum, mResultOwner, mResultText, mResultOptSeq, mPureResults, mSortedResults, mBestResult, mAvgResult, mIsBestBr, mIsAvgBr

    Private Sub Class_Initialize
        mPostNum = ""
        mResultOwner = ""
        mResultText = ""
        mResultOptSeq = ""
        mPureResults = ""
        mSortedResults = ""
        mBestResult = ""
        mAvgResult = ""
        mIsBestBr = False
        mIsAvgBr = False
    End Sub

    Public Property Let PostNum(value)
        mPostNum = value
    End Property

    Public Property Let ResultOwner(value)
        mResultOwner = value
    End Property

    Public Property Let ResultText(value)
        mResultText = value
    End Property

    Public Property Let ResultOptSeq(value)
        mResultOptSeq = value
    End Property

    Public Property Let PureResults(value)
        mPureResults = value
    End Property

    Public Property Let SortedResults(value)
        mSortedResults = value
    End Property

    Public Property Let BestResult(value)
        mBestResult = value
    End Property

    Public Property Let AvgResult(value)
        mAvgResult = value
    End Property

    Public Property Let IsBestBr(value)
        mIsBestBr = value
    End Property

    Public Property Let IsAvgBr(value)
        mIsAvgBr = value
    End Property

    Public Property Get PostNum
        PostNum = mPostNum
    End Property

    Public Property Get ResultOwner
        ResultOwner = mResultOwner
    End Property

    Public Property Get ResultText
        ResultText = mResultText
    End Property

    Public Property Get ResultOptSeq
        ResultOptSeq = mResultOptSeq
    End Property

    Public Property Get PureResults
        PureResults = mPureResults
    End Property

    Public Property Get SortedResults
        SortedResults = mSortedResults
    End Property

    Public Property Get BestResult
        BestResult = mBestResult
    End Property

    Public Property Get AvgResult
        AvgResult = mAvgResult
    End Property

    Public Property Get IsBestBr
        IsBestBr = mIsBestBr
    End Property

    Public Property Get IsAvgBr
        IsAvgBr = mIsAvgBr
    End Property
End Class



Class InvalidResultInfo
    Private mPostNum, mResultOwner, mResultText, mResultOptSeq, mPureResults

    Private Sub Class_Initialize
        mPostNum = ""
        mResultOwner = ""
        mResultText = ""
        mResultOptSeq = ""
        mPureResults = ""
    End Sub

    Public Property Let PostNum(value)
        mPostNum = value
    End Property

    Public Property Let ResultOwner(value)
        mResultOwner = value
    End Property

    Public Property Let ResultText(value)
        mResultText = value
    End Property

    Public Property Let ResultOptSeq(value)
        mResultOptSeq = value
    End Property

    Public Property Let PureResults(value)
        mPureResults = value
    End Property

    Public Property Get PostNum
        PostNum = mPostNum
    End Property

    Public Property Get ResultOwner
        ResultOwner = mResultOwner
    End Property

    Public Property Get ResultText
        ResultText = mResultText
    End Property

    Public Property Get ResultOptSeq
        ResultOptSeq = mResultOptSeq
    End Property

    Public Property Get PureResults
        PureResults = mPureResults
    End Property
End Class



Class VariableArray
    Private mPreBound, mBound, mArray()

    Private Sub Class_Initialize
        mPreBound = -1
        mBound = -1
    End Sub

    Public Sub SetPreBound(sValue)
        If isNumeric(sValue) Then
            If sValue > mBound Then
                ReDim Preserve mArray(sValue)
                mPreBound = sValue
            End If
        End If
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
        If mBound > mPreBound Then
            ReDim Preserve mArray(mBound)
            mPreBound = mBound
        End If

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

        Call ResetArray()
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
            If isObject(mArray(seq)) Then
                For i = seq To mBound - 1
                    Set mArray(i) = mArray(i + 1)
                Next
            Else
                For i = seq To mBound - 1
                    mArray(i) = mArray(i + 1)
                Next
            End If
        End If

        mBound = mBound - 1
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

    Public Function ToStringWithSpace()
        If mBound = -1 Then ToStringWithSpace = "" : Exit Function

        Dim i, sTmp
        For i = 0 To mBound
            sTmp = sTmp & " " & mArray(i)
        Next

        ToStringWithSpace = Trim(sTmp)
    End Function

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