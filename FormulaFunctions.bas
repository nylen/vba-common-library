Attribute VB_Name = "FormulaFunctions"
' Common VBA Library
' FormulaFunctions
' Provides functions that are useful in Excel formulas.

Option Explicit

' Retrieves the given element of an array.
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

' Splits a string into an array, optionally limiting the number
' of items in the returned array.
Public Function StringSplit(s As String, delim As String, _
    Optional limit As Long = -1) As String()
    
    StringSplit = Split(s, delim, limit)
End Function

' Returns a newline (vbLf) character for use in formulas.
Public Function NewLine() As String
    NewLine = vbLf
End Function

' Returns an array suitable for using in an array formula.  When this
' function is called from an array formula, it will detect whether or
' not the array should be transposed to fit into the range.
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
                Err.Raise 32000, Description:= _
                    "Invalid number of dimensions (" & Rank(arr) _
                        & "; expected 1 or 2)."
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

' Returns the width of a column on a sheet.  If the column number is
' not given and this function is used in a formula, then it returns
' the column width of the cell containing the formula.
Public Function ColumnWidth(Optional c As Integer = 0) As Variant
    Application.Volatile
    Dim s As Worksheet
    If IsObject(Application.Caller) Then
        Set s = Application.Caller.Worksheet
    Else
        Set s = ActiveSheet
    End If
    If c <= 0 And IsObject(Application.Caller) Then
        c = Application.Caller.Column
    End If
    ColumnWidth = s.Columns(c).Width
End Function

' Returns the height of a row on a sheet.  If the row number is
' not given and this function is used in a formula, then it returns
' the row height of the cell containing the formula.
Public Function RowHeight(Optional r As Integer = 0) As Variant
    Application.Volatile
    Dim s As Worksheet
    If IsObject(Application.Caller) Then
        Set s = Application.Caller.Worksheet
    Else
        Set s = ActiveSheet
    End If
    If r <= 0 And IsObject(Application.Caller) Then
        r = Application.Caller.Row
    End If
    RowHeight = s.Rows(r).Height
End Function

