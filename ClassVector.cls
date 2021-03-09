VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================================
' Author:   Vladimir Dmitriev <v.a.dmitriev123@gmail.com>
' Link:     https://github.com/vadmitriev/
' =========================================================

Private n_arr       As Variant
Private n_showError As Boolean
Private n_extractArr As Boolean
Private Const START_VALUE = vbNullString

Private Sub Class_initialize()
    ReDim n_arr(0)
    n_arr(0) = START_VALUE
    n_showError = True
    n_extractArr = True
End Sub

Public Function Add(ParamArray Args() As Variant)
    Call addItemsFromArgsToArr(n_arr, Args)
End Function

Private Sub addItemsFromArgsToArr(ByRef arr As Variant, ParamArray Args() As Variant)
    Dim i           As Long
    Dim j           As Long
    Dim newVal      As Variant
    
    For i = LBound(Args) To UBound(Args)
        If VBA.IsObject(Args(i)) Then
            Set newVal = Args(i)
        Else
            newVal = Args(i)
        End If
        If IsArray(newVal) Then
            For j = LBound(newVal) To UBound(newVal)
                Call addItemsFromArgsToArr(arr, newVal(j))
            Next j
        Else
            Call pushValueToArr(arr, newVal)
        End If
    Next i
End Sub

Private Sub pushValueToArr(ByRef arr As Variant, ByRef newVal As Variant)
    If VarType(arr(0)) = VarType(START_VALUE) Then
        If arr(0) = START_VALUE Then
            arr(0) = newVal
        Else
            ReDim Preserve arr(UBound(arr) + 1)
            arr(UBound(arr)) = newVal
        End If
    Else
        ReDim Preserve arr(UBound(arr) + 1)
        If VBA.IsObject(newVal) Then
            Set arr(UBound(arr)) = newVal
        Else
            arr(UBound(arr)) = newVal
        End If
    End If
End Sub

Public Sub Clear()
    ReDim n_arr(0)
End Sub

Public Sub Concat(ParamArray Args() As Variant)
    Dim i           As Long
    Dim item_       As Variant
    
    For i = LBound(Args) To UBound(Args)
        Set item_ = Args(i)
        If TypeName(item_) = TypeName(Me) Then
            Me.Add (item_.values)
        End If
    Next i
End Sub

Public Sub Convert(ByVal arr As Variant)
    n_arr = arr
End Sub

Public Function Copy() As ClassVector
    Dim newVector   As New ClassVector
    
    Call newVector.Add(n_arr)

    Set Copy = newVector
End Function

Public Function Count() As Long
    If n_arr(0) = START_VALUE Then
        Count = 0
    Else
        Count = UBound(n_arr) + 1
    End If
End Function

Public Sub Delete(ParamArray Args() As Variant)
    Call Me.DeleteByValue(Args)
End Sub

Public Sub DeleteByIndex(ParamArray Args() As Variant)
    Dim i           As Long
    Dim j           As Long
    Dim item_       As Variant
    Dim newArr      As Variant
    Dim dictIndex
    Set dictIndex = CreateObject("Scripting.Dictionary")
    
    ReDim newArr(0)
    newArr(0) = START_VALUE
    For i = LBound(Args) To UBound(Args)
        dictIndex(Args(i)) = i
    Next i
    
    For i = LBound(n_arr) To UBound(n_arr)
        If Not dictIndex.exists(i) Then
            item_ = n_arr(i)
            Call pushValueToArr(newArr, item_)
        End If
    Next i
    
    n_arr = newArr
End Sub

Public Sub DeleteByValue(ParamArray Args() As Variant)
    Dim i           As Long
    Dim j           As Long
    Dim item_       As Variant
    Dim newArr      As Variant
    Dim dictValues
    Set dictValues = CreateObject("Scripting.Dictionary")
    
    ReDim newArr(0)
    newArr(0) = START_VALUE
    For i = LBound(Args) To UBound(Args)
        item_ = Args(i)
        If IsArray(item_) Then
            For j = LBound(item_) To UBound(item_)
                dictValues(item_(j)) = j
            Next j
        Else
            dictValues(Args(i)) = i
        End If
    Next i
    
    For i = LBound(n_arr) To UBound(n_arr)
        item_ = n_arr(i)
        If Not dictValues.exists(item_) Then
            Call pushValueToArr(newArr, item_)
        End If
    Next i
    
    n_arr = newArr
End Sub

Public Function Exist(ParamArray Args() As Variant) As Boolean
    Dim i           As Long
    Dim j           As Long
    Dim item_       As Variant
    Dim valTest     As Variant
    Dim Count       As Long
    
    Count = 0
    For i = LBound(n_arr) To UBound(n_arr)
        item_ = n_arr(i)
        For j = LBound(Args) To UBound(Args)
            valTest = Args(j)
            If item_ = valTest Then
                Count = Count + 1
            End If
        Next j
    Next i
    
    Exist = (Count = UBound(Args) + 1)
End Function

Public Function FilterType(ByRef needType As Variant, Optional excludeThis As Boolean = False) As Variant
    Dim i           As Long
    Dim item_       As Variant
    Dim newArr      As Variant
    Dim checkType   As Boolean
    Dim needAdd     As Boolean
    
    ReDim newArr(0)
    newArr(0) = START_VALUE
    For i = LBound(n_arr) To UBound(n_arr)
        item_ = n_arr(i)
        checkType = UCase(Trim(TypeName(item_))) = UCase(Trim(needType)) Or _
                    VarType(item_) = needType
        needAdd = IIf(excludeThis, Not checkType, checkType)
        If needAdd Then Call pushValueToArr(newArr, item_)
    Next i
    
    FilterType = newArr
End Function

Public Function First() As Variant
    If Me.isEmpty() Then
        First = vbEmpty
    Else
        First = n_arr(LBound(n_arr))
    End If
End Function

Public Function GetUniqueValues(Optional includeArrs As Boolean = True) As Variant
    GetUniqueValues = spreadArrUnique(n_arr)
End Function

Private Function spreadArrUnique(ByVal arr As Variant) As Variant
    Dim i           As Long
    Dim j           As Long
    Dim item_       As Variant
    Dim dict
    Set dict = CreateObject("Scripting.Dictionary")

    For i = LBound(arr) To UBound(arr)
        item_ = arr(i)
        If IsArray(item_) Then
            item_ = spreadArrUnique(item_)
            For j = LBound(item_) To UBound(item_)
                dict.item(item_(j)) = j
            Next j
        Else
            If item_ <> vbEmpty Then
                dict.item(item_) = i
            End If
        End If
    Next i

    spreadArrUnique = dict.keys()
End Function

Public Function GetValuesByIndex(ParamArray Args() As Variant) As Variant
    Dim i           As Long
    Dim j           As Long
    Dim item_       As Variant
    Dim newArr      As Variant
    Dim onlyOne     As Boolean
    Dim dictIndex
    Set dictIndex = CreateObject("Scripting.Dictionary")
    
    ReDim newArr(0)
    newArr(0) = START_VALUE
    For i = LBound(Args) To UBound(Args)
        dictIndex(Args(i)) = i
    Next i
    
    For i = LBound(n_arr) To UBound(n_arr)
        If dictIndex.exists(i) Then
            item_ = n_arr(i)
            Call pushValueToArr(newArr, item_)
        End If
    Next i
    
    If UBound(newArr) = 0 And Not IsArray(newArr(0)) Then
        GetValuesByIndex = CStr(newArr(0))
    Else
        GetValuesByIndex = newArr
    End If
End Function

Public Sub InsertAfter(ByVal index As Long, ParamArray Args() As Variant)
    Call insertAfterIndex(index, Args)
End Sub

Public Sub InsertBefore(ByVal index As Long, ParamArray Args() As Variant)
    Call insertAfterIndex(index - 1, Args)
End Sub

Private Sub insertAfterIndex(ByVal index As Long, ParamArray Args() As Variant)
    Dim newArr      As Variant
    Dim newIndex    As Long
    Dim needAdd     As Boolean
    Dim i           As Long
    
    newIndex = checkBounds(index)
    ReDim newArr(0)
    newArr(0) = START_VALUE
    For i = LBound(n_arr) To UBound(n_arr)
        needAdd = True
        If newIndex = UBound(n_arr) Then
            Call pushValueToArr(newArr, n_arr(i))
            needAdd = False
        End If
        If i = newIndex Then
            If needAdd And newIndex <> LBound(n_arr) Then
                Call pushValueToArr(newArr, n_arr(i))
                needAdd = False
            End If
            Call addItemsFromArgsToArr(newArr, Args)
        End If
        If needAdd Then
            Call pushValueToArr(newArr, n_arr(i))
        End If
    Next i

    n_arr = newArr
End Sub

Public Function isEmpty() As Boolean
    isEmpty = (n_arr(0) = START_VALUE)
End Function

Private Function checkBounds(ByVal index As Variant) As Long
    If index < LBound(n_arr) Then
        If Me.showError Then
            Err.Raise Number:=vbObjectError + 513, _
                Description:="Error! Required index (" & index & ") less then low bound of array"
        End If
        checkBounds = LBound(n_arr)
        Exit Function
    End If
    
    If index > UBound(n_arr) Then
        If Me.showError Then
            Err.Raise Number:=vbObjectError + 513, _
                    Description:="Error! Required index (" & index & ") more then up bound of array"
        End If
        checkBounds = UBound(n_arr)
        Exit Function
    End If
    
    checkBounds = index
End Function

Public Function Last() As Variant
    If Me.isEmpty() Then
        Last = vbEmpty
    Else
        Last = n_arr(UBound(n_arr))
    End If
End Function

Public Sub ReplaceString(ByVal oldValue As Variant, ByVal newValue As Variant, _
                            Optional replaceSubstring As Boolean = True, _
                            Optional matchCase As Boolean = False)
    Dim i           As Long
    Dim j           As Long
    Dim item_       As Variant
    Dim testRepl    As Boolean
    Dim compareMethod As VbCompareMethod
    
    For i = LBound(n_arr) To UBound(n_arr)
        item_ = n_arr(i)
        If VarType(item_) = vbString Then
            If replaceSubstring Then
                compareMethod = IIf(matchCase, vbBinaryCompare, vbTextCompare)
                n_arr(i) = Replace(item_, oldValue, newValue, compare:=compareMethod)
            Else
                testRepl = False
                If matchCase Then
                    If item_ = newStr Then testRepl = True
                Else
                    If Trim(UCase(item_)) = Trim(UCase(newValue)) Then testRepl = True
                End If
                If testRepl Then n_arr(i) = newValue
            End If
        End If
    Next i
End Sub

Public Sub ReplaceValue(ByVal oldValue As Variant, ByVal newValue As Variant, Optional replaceInArr As Boolean = False)
    Dim i           As Long
    Dim j           As Long
    Dim item_       As Variant
    
    For i = LBound(n_arr) To UBound(n_arr)
        item_ = n_arr(i)
        If IsArray(item_) And Not IsArray(newValue) Then
            If replaceInArr Then
                For j = LBound(item_) To UBound(item_)
                    If item_(j) = oldValue Then
                        item_(j) = newValue
                        n_arr(i) = item_
                    End If
                Next j
            End If
        Else
            If item_ = oldValue Then
                n_arr(i) = newValue
            End If
        End If
    Next i
End Sub

Public Function Slice(ByVal beginIndex As Long, Optional ByVal endIndex = -1) As Variant
    Dim newArr      As Variant
    Dim newBeginIndex As Long
    Dim countVals   As Long
    Dim i           As Long

    newBeginIndex = checkBounds(beginIndex)
    If endIndex > UBound(n_arr) Or endIndex = -1 Then endIndex = UBound(n_arr)
    
    ReDim newArr(endIndex - newBeginIndex)
    countVals = 0
    For i = newBeginIndex To endIndex
        newArr(countVals) = n_arr(i)
        countVals = countVals + 1
    Next i
    
    Slice = newArr
End Function

Public Sub Sort(Optional lowToHigh As Boolean = True)
    Call QuickSort(n_arr, LBound(n_arr), UBound(n_arr))
    
    Dim newArr      As Variant
    Dim i           As Long
    Dim Count       As Long
    ReDim newArr(UBound(n_arr))
    If Not lowToHigh Then
        Count = 0
        For i = UBound(n_arr) To LBound(n_arr) Step -1
            newArr(Count) = n_arr(i)
            Count = Count + 1
        Next i
        n_arr = newArr
    End If
End Sub

Public Function SumValues(Optional includeArrs As Boolean = True) As Double
    Dim sumValue    As Double
    Dim i           As Long
    Dim j           As Long
    Dim item        As Variant
    
    sumValue = 0
    
    For i = LBound(n_arr) To UBound(n_arr)
        item = n_arr(i)
        If IsArray(item) Then
            If includeArrs Then sumValue = sumValue + sumInArr(item)
        Else
            If IsNumeric(item) Then sumValue = sumValue + item
        End If
    Next i
    
    SumValues = sumValue
End Function

Private Function sumInArr(ByRef arr As Variant) As Double
    Dim i           As Long
    Dim sum_        As Double
    Dim item        As Variant
    
    sum_ = 0
    For i = LBound(arr) To UBound(arr)
        item = arr(i)
        If IsArray(item) Then
            sum_ = sumInArr(item)
        Else
            If IsNumeric(item) Then
                sum_ = sum_ + item
            End If
        End If
    Next i
    
    sumInArr = sum_
End Function

Public Function toString(Optional prefix As String = ",", Optional includeArrs As Boolean = True) As String
    Dim line        As String
    Dim newArr      As Variant
    Dim item        As Variant
    Dim i           As Long
    
    line = vbNullString
    For i = LBound(n_arr) To UBound(n_arr)
        If VBA.IsObject(n_arr(i)) Then
            Set item = n_arr(i)
        Else
            item = n_arr(i)
        End If
        If IsArray(item) Then
            If includeArrs Then line = line & prefix & Join(item, prefix)
        Else
            If Not VBA.IsObject(item) Then
                line = line & prefix & item
            End If
        End If
    Next i
    
    If Left(line, Len(prefix)) = prefix Then line = Right(line, Len(line) - Len(prefix))
    toString = line
End Function

Private Sub QuickSort(ByRef vArray As Variant, inLow As Long, inHi As Long)
    ' Link: https://stackoverflow.com/a/152325
    
    Dim pivot   As Variant
    Dim tmpSwap As Variant
    Dim tmpLow  As Long
    Dim tmpHi   As Long
    
    tmpLow = inLow
    tmpHi = inHi
    
    pivot = vArray((inLow + inHi) \ 2)
    
    While (tmpLow <= tmpHi)
        While (vArray(tmpLow) < pivot And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend
        
        While (pivot < vArray(tmpHi) And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend
        
        If (tmpLow <= tmpHi) Then
            tmpSwap = vArray(tmpLow)
            vArray(tmpLow) = vArray(tmpHi)
            vArray(tmpHi) = tmpSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Wend
    
    If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
    If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
End Sub

Public Property Get showError() As Boolean
    showError = n_showError
End Property

Public Property Let showError(newVal As Boolean)
    n_showError = newVal
End Property

Public Property Get extractArrayToAdd() As Boolean
    extractArrayToAdd = n_extractArr
End Property

Public Property Let extractArrayToAdd(newVal As Boolean)
    n_extractArr = newVal
End Property

Public Property Get values() As Variant
    Dim newArr      As Variant
    ReDim newArr(0)
    
    If Me.isEmpty() Then
        values = newArr
    Else
        values = n_arr
    End If
End Property

Public Property Set values(arr As Variant)
    If Not IsArray(arr) Then
        n_arr = Array(arr)
    Else
        n_arr = arr
    End If
End Property

Public Property Get item(index As Variant) As Variant
    Dim newIndex    As Long

    newIndex = checkBounds(index)
    If VBA.IsObject(n_arr(newIndex)) Then
        Set item = n_arr(newIndex)
    Else
        item = n_arr(newIndex)
    End If
End Property

Public Property Set item(index As Variant, ByRef newValue)
    Dim newIndex    As Long
    newIndex = checkBounds(index)
    
    If VBA.IsObject(n_arr(newIndex)) Then
        Set n_arr(newIndex) = newValue
    Else
        n_arr(newIndex) = newValue
    End If
End Property
