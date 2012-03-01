Attribute VB_Name = "StringUtils"
Option Explicit

Public Function NewLine() As String
    NewLine = vbLf
End Function

Public Function StartsWith(s As String, prefix As String, _
    Optional caseSensitive As Boolean = True) As Boolean
    
    If caseSensitive Then
        StartsWith = (Left(s, Len(prefix)) = prefix)
    Else
        StartsWith = (Left(LCase(s), Len(prefix)) = LCase(prefix))
    End If
End Function

Public Function EndsWith(s As String, suffix As String, _
    Optional caseSensitive As Boolean = True) As Boolean
    
    If caseSensitive Then
        EndsWith = (Right(s, Len(suffix)) = suffix)
    Else
        EndsWith = (Right(LCase(s), Len(suffix)) = LCase(suffix))
    End If
End Function

Public Function SplitTrim(s As String, delim As String) As String()
    Dim arr() As String
    arr = Split(s, delim)
    
    Dim i As Integer
    For i = 0 To UBound(arr)
        arr(i) = Trim(arr(i))
    Next
    
    SplitTrim = arr
End Function

Public Function TrimChars(s As String, toTrim As String)
    TrimChars = TrimTrailingChars(TrimLeadingChars(s, toTrim), toTrim)
End Function

Public Function TrimLeadingChars(s As String, toTrim As String)
    If s = "" Then
        TrimLeadingChars = ""
        Exit Function
    End If
    Dim i As Integer
    i = 1
    While InStr(toTrim, Mid(s, i, 1)) > 0 And i <= Len(s)
        i = i + 1
    Wend
    TrimLeadingChars = Mid(s, i)
End Function

Public Function TrimTrailingChars(s As String, toTrim As String)
    If s = "" Then
        TrimTrailingChars = ""
        Exit Function
    End If
    Dim i As Integer
    i = Len(s)
    While InStr(toTrim, Mid(s, i, 1)) > 0 And i >= 1
        i = i - 1
    Wend
    TrimTrailingChars = Left(s, i)
End Function
