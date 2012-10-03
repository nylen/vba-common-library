Attribute VB_Name = "VBALib_ArrayUtils"
' Common VBA Library, version 2012-10-03.1
' ArrayUtils
' Provides functions for handling arrays that are lacking in the built-in VBA
' language.

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (dest As Any, source As Any, ByVal bytes As Long)

Private Const NORMALIZE_LBOUND = 1

' Returns a single-dimension array with lower bound 1, if given a
' one-dimensional array with any lower bound or a two-dimensional array with
' with one dimension having only one element.  This function will always
' return a copy of the given array.
Public Function NormalizeArray(arr As Variant) As Variant
    If ArrayLen(arr) = 0 Then
        NormalizeArray = Array()
        Exit Function
    End If
    
    Dim arr2() As Variant
    
    Dim nItems As Long
    Dim i As Long
    
    Select Case Rank(arr)
        Case 1
            If LBound(arr) = NORMALIZE_LBOUND Then
                NormalizeArray = arr
            Else
                nItems = ArrayLen(arr)
                ReDim arr2(NORMALIZE_LBOUND To NORMALIZE_LBOUND + nItems - 1)
                For i = NORMALIZE_LBOUND To NORMALIZE_LBOUND + nItems - 1
                    arr2(i) = arr(i + LBound(arr) - NORMALIZE_LBOUND)
                Next
                NormalizeArray = arr2
            End If
            
        Case 2
            If LBound(arr, 1) = UBound(arr, 1) Then
                
                ' Copy values from array's second dimension
                nItems = ArrayLen(arr, 2)
                ReDim arr2(NORMALIZE_LBOUND To NORMALIZE_LBOUND + nItems - 1)
                For i = NORMALIZE_LBOUND To NORMALIZE_LBOUND + nItems - 1
                    arr2(i) = arr(LBound(arr, 1), _
                        i + LBound(arr, 2) - NORMALIZE_LBOUND)
                Next
                NormalizeArray = arr2
                
            ElseIf LBound(arr, 2) = UBound(arr, 2) Then
                
                ' Copy values from array's first dimension
                nItems = ArrayLen(arr, 1)
                ReDim arr2(NORMALIZE_LBOUND To NORMALIZE_LBOUND + nItems - 1)
                For i = NORMALIZE_LBOUND To NORMALIZE_LBOUND + nItems - 1
                    arr2(i) = arr(i + LBound(arr, 1) - NORMALIZE_LBOUND, _
                        LBound(arr, 2))
                Next
                NormalizeArray = arr2
                
            Else
                Err.Raise 32000, Description:= _
                    "Can only normalize a 2-dimensional array if one of " _
                        & "the dimensions contains only one element."
            End If
            
        Case Else
            Err.Raise 32000, Description:= _
                "Can only normalize 1- and 2-dimensional arrays."
    End Select
End Function

' Returns the rank (number of dimensions) of an array.
' From http://www.devx.com/vb2themax/Tip/18265 .
Public Function Rank(arr As Variant) As Integer
    Dim ptr As Long
    Dim vType As Integer
    Const VT_BYREF = &H4000&
    
    ' get the real VarType of the argument
    ' this is similar to VarType(), but returns also the VT_BYREF bit
    CopyMemory vType, arr, 2
    
    ' exit if not an array
    If (vType And vbArray) = 0 Then Exit Function
    
    ' get the address of the SAFEARRAY descriptor
    ' this is stored in the second half of the
    ' Variant parameter that has received the array
    CopyMemory ptr, ByVal VarPtr(arr) + 8, 4
    
    ' see whether the routine was passed a Variant
    ' that contains an array, rather than directly an array
    ' in the former case ptr already points to the SA structure.
    ' Thanks to Monte Hansen for this fix
    If (vType And VT_BYREF) Then
        ' ptr is a pointer to a pointer
        CopyMemory ptr, ByVal ptr, 4
    End If
    
    ' get the address of the SAFEARRAY structure
    ' this is stored in the descriptor
    ' get the first word of the SAFEARRAY structure
    ' which holds the number of dimensions
    ' ...but first check that saAddr is non-zero, otherwise
    ' this routine bombs when the array is uninitialized
    ' (Thanks to VB2TheMax aficionado Thomas Eyde for
    ' suggesting this edit to the original routine.)
    If ptr Then
        CopyMemory Rank, ByVal ptr, 2
    End If
End Function

' Returns the number of elements in an array for a given dimension.
Public Function ArrayLen(arr As Variant, _
    Optional dimNum As Integer = 1) As Long
    
    If IsEmpty(arr) Then
        ArrayLen = 0
    Else
        ArrayLen = UBound(arr, dimNum) - LBound(arr, dimNum) + 1
    End If
End Function

' Sorts a section of an array in place.  Code from:
' http://stackoverflow.com/questions/152319/vba-array-sort-function
Private Sub QuickSort(vArray() As Variant, inLow As Long, inHi As Long)
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

' Sorts the given single-dimension array in place.
Public Sub SortArrayInPlace(arr() As Variant)
    QuickSort arr, LBound(arr), UBound(arr)
End Sub

' Returns a sorted copy of the given array.
Public Function SortArray(arr() As Variant) As Variant()
    If ArrayLen(arr) = 0 Then
        SortArray = Array()
    Else
        Dim arr2() As Variant
        arr2 = arr
        SortArrayInPlace arr2
        SortArray = arr2
    End If
End Function

' Returns an array containing each unique item in the given array only once.
Public Function GetUniqueItems(arr() As Variant) As Variant()
    If ArrayLen(arr) = 0 Then
        GetUniqueItems = Array()
    Else
        Dim arrSorted() As Variant
        arrSorted = SortArray(arr)
        
        Dim uniqueItemsList As New VBALib_List
        uniqueItemsList.Add arrSorted(LBound(arrSorted))
        
        Dim i As Long
        For i = LBound(arrSorted) + 1 To UBound(arrSorted)
            If arrSorted(i) <> arrSorted(i - 1) Then
                uniqueItemsList.Add arrSorted(i)
            End If
        Next
        
        GetUniqueItems = uniqueItemsList.Items
    End If
End Function
