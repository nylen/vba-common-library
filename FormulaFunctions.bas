Attribute VB_Name = "FormulaFunctions"
Option Explicit

Public Function ArrayElement(arr As Variant, i1 As Variant, _
    Optional i2 As Variant, Optional i3 As Variant, _
    Optional i4 As Variant, Optional i5 As Variant) As Variant
    
    If IsMissing(i2) Then
        ArrayElement = arr(i1)
    ElseIf IsMissing(i3) Then
        ArrayElement = arr(i1, i2)
    ElseIf IsMissing(i4) Then
        ArrayElement = arr(i1, i2, i3)
    ElseIf IsMissing(i5) Then
        ArrayElement = arr(i1, i2, i3, i4)
    Else
        ArrayElement = arr(i1, i2, i3, i4, i5)
    End If
End Function

Public Function StringSplit(s As String, delim As String, _
    Optional limit As Long = -1) As String()
    
    StringSplit = Split(s, delim, limit)
End Function

Public Function RangeArray(arr As Variant) As Variant
    If IsObject(Application.Caller) Then
        Dim len1 As Long, len2 As Long
        Select Case Rank(arr)
            Case 0
                RangeArray = Empty
                Exit Function
            Case 1
                len1 = ArrayLen(arr)
                len2 = 1
            Case 2
                len1 = ArrayLen(arr)
                len2 = ArrayLen(arr, 2)
            Case Else
                Err.Raise 32000, _
                    Description:="Invalid number of dimensions (" _
                        & Rank(arr) & "; expected 1 or 2)."
        End Select
        
        If Application.Caller.Rows.Count > Application.Caller.Columns.Count _
            And len1 > len2 Then
            
            RangeArray = WorksheetFunction.Transpose(arr)
        Else
            RangeArray = arr
        End If
    Else
        RangeArray = arr
    End If
End Function
