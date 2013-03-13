Attribute VB_Name = "VBALib_StringUtils"
' Common VBA Library - StringUtils
' Provides useful functions for manipulating strings.

Option Explicit

' Determines whether a string starts with a given prefix.
Public Function StartsWith(s As String, prefix As String, _
    Optional caseSensitive As Boolean = True) As Boolean
    
    If caseSensitive Then
        StartsWith = (Left(s, Len(prefix)) = prefix)
    Else
        StartsWith = (Left(LCase(s), Len(prefix)) = LCase(prefix))
    End If
End Function

' Determines whether a string ends with a given suffix.
Public Function EndsWith(s As String, suffix As String, _
    Optional caseSensitive As Boolean = True) As Boolean
    
    If caseSensitive Then
        EndsWith = (Right(s, Len(suffix)) = suffix)
    Else
        EndsWith = (Right(LCase(s), Len(suffix)) = LCase(suffix))
    End If
End Function

' Splits a string on a given delimiter, trimming trailing and leading
' whitespace from each piece of the string.
Public Function SplitTrim(s As String, delim As String) As String()
    Dim arr() As String
    arr = Split(s, delim)
    
    Dim i As Integer
    For i = 0 To UBound(arr)
        arr(i) = Trim(arr(i))
    Next
    
    SplitTrim = arr
End Function

' Trims a specified set of characters from the beginning and end
' of the given string.
' @param toTrim: The characters to trim.  For example, if ",; "
' is given, then all spaces, commas, and semicolons will be removed
' from the beginning and end of the given string.
Public Function TrimChars(s As String, toTrim As String)
    TrimChars = TrimTrailingChars(TrimLeadingChars(s, toTrim), toTrim)
End Function

' Trims a specified set of characters from the beginning of the
' given string.
' @param toTrim: The characters to trim.  For example, if ",; "
' is given, then all spaces, commas, and semicolons will be removed
' from the beginning of the given string.
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

' Trims a specified set of characters from the end of the given
' string.
' @param toTrim: The characters to trim.  For example, if ",; "
' is given, then all spaces, commas, and semicolons will be removed
' from the end of the given string.
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
