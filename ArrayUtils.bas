Attribute VB_Name = "ArrayUtils"
' Common VBA Library
' ArrayUtils
' Provides functions for handling arrays that are lacking in the built-in VBA language.

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (dest As Any, source As Any, ByVal bytes As Long)

Private Const NORMALIZE_LBOUND = 1

' Returns a single-dimension array with lower bound 1, if given
' a one dimensional array with any lower bound or a two-dimensional
' array with one dimension having only one element.
Public Function NormalizeArray(arr As Variant) As Variant
    If IsEmpty(arr) Then
        NormalizeArray = Empty
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
                Err.Raise 32000, _
                    Description:="Can only normalize a 2-dimensional array " _
                        & "if one of the dimensions contains only one element."
            End If
            
        Case Else
            Err.Raise 32000, _
                Description:="Can only normalize 1- and 2-dimensional arrays."
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
Function ArrayLen(arr As Variant, Optional dimNum As Integer = 1) As Long
    If IsEmpty(arr) Then
        ArrayLen = 0
    Else
        ArrayLen = UBound(arr, dimNum) - LBound(arr, dimNum) + 1
    End If
End Function
