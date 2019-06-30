Attribute VB_Name = "UtilityArray"
Option Explicit

Public Function Range(n As Long) As Variant
    
    Dim i As Long
    
    Dim output() As Long
    ReDim output(1 To n)
    
    For i = 1 To n
        output(i) = i
    Next
    
    Range = output
    
End Function

Public Function RepVal(value As Variant, n As Long) As Variant
    
    Dim i As Long
    
    Dim output() As Variant
    ReDim output(1 To n)
    
    For i = 1 To n
        output(i) = value
    Next
    
    RepVal = output
    
End Function

Public Function AddArr(inputArray As Variant, otherArray As Variant) As Variant
    
    Dim i As Long
    
    Dim output() As Variant
    ReDim output(LBound(inputArray) To UBound(inputArray))
    
    For i = LBound(inputArray) To UBound(inputArray)
        output(i) = inputArray(i) + otherArray(i)
    Next
    
    AddArr = output
    
End Function

Public Function MulArr(inputArray As Variant, otherArray As Variant) As Variant
    
    Dim i As Long
    
    Dim output() As Variant
    ReDim output(LBound(inputArray) To UBound(inputArray))
    
    For i = LBound(inputArray) To UBound(inputArray)
        output(i) = inputArray(i) * otherArray(i)
    Next
    
    MulArr = output
    
End Function

Public Function AddVal(inputArray As Variant, value As Variant) As Variant
    
    Dim i As Long
    
    Dim output() As Variant
    ReDim output(LBound(inputArray) To UBound(inputArray))
    
    For i = LBound(inputArray) To UBound(inputArray)
        output(i) = inputArray(i) + value
    Next
    
    AddVal = output
    
End Function

Public Function MulVal(inputArray As Variant, value As Variant) As Variant
    
    Dim i As Long
    
    Dim output() As Variant
    ReDim output(LBound(inputArray) To UBound(inputArray))
    
    For i = LBound(inputArray) To UBound(inputArray)
        output(i) = inputArray(i) * value
    Next
    
    MulVal = output
    
End Function

Public Sub SwapVal(A As Variant, B As Variant)
    
    Dim tmpSwap As Variant
    
    tmpSwap = A
    A = B
    B = tmpSwap
    
End Sub

Public Sub SwapObj(A As Variant, B As Variant)
    
    Dim tmpSwap As Variant
    
    Set tmpSwap = A
    Set A = B
    Set B = tmpSwap
    
End Sub

Public Sub ReverseVal(inputArray As Variant)
    
    Dim i As Long
    
    Dim n As Long
    n = UBound(inputArray)
    
    For i = LBound(inputArray) To ((UBound(inputArray) - LBound(inputArray) + 1) \ 2)
        SwapVal inputArray(i), inputArray(n)
        n = n - 1
    Next
    
End Sub

Public Sub ReverseObj(inputArray As Variant)
    
    Dim i As Long
    
    Dim n As Long
    n = UBound(inputArray)
    
    For i = LBound(inputArray) To ((UBound(inputArray) - LBound(inputArray) + 1) \ 2)
        SwapObj inputArray(i), inputArray(n)
        n = n - 1
    Next
    
End Sub

Public Sub QuickSortObjectByValue(objectArray As Variant, valueArray As Variant, inLo As Long, inHi As Long)
    
    Dim pivot As Variant
    Dim tmpLo As Long
    Dim tmpHi As Long
    
    tmpLo = inLo
    tmpHi = inHi
    
    pivot = valueArray((inLo + inHi) \ 2)
    
    While (tmpLo <= tmpHi)
        While (pivot > valueArray(tmpLo) And tmpLo < inHi)
            tmpLo = tmpLo + 1
        Wend
        
        While (pivot < valueArray(tmpHi) And inLo < tmpHi)
            tmpHi = tmpHi - 1
        Wend
        
        If (tmpLo <= tmpHi) Then
            SwapVal valueArray(tmpLo), valueArray(tmpHi)
            SwapObj objectArray(tmpLo), objectArray(tmpHi)
            tmpLo = tmpLo + 1
            tmpHi = tmpHi - 1
        End If
    Wend
    
    If (inLo < tmpHi) Then QuickSortObjectByValue objectArray, valueArray, inLo, tmpHi
    If (tmpLo < inHi) Then QuickSortObjectByValue objectArray, valueArray, tmpLo, inHi
    
End Sub
