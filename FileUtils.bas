Attribute VB_Name = "FileUtils"
Option Explicit

Public Function FileExists(ByVal testFilename As String, _
    Optional findFolders As Boolean = False) As Boolean
    
    ' Include read-only files, hidden files, system files.
    Dim attrs As Long
    attrs = (vbReadOnly Or vbHidden Or vbSystem)
    
    If findFolders Then
        attrs = (attrs Or vbDirectory) ' Include folders as well.
    End If
    
    'If Dir() returns something, the file exists.
    On Error Resume Next
    FileExists = (Dir(TrimTrailingChars(testFilename, "/\"), attrs) <> "")
End Function

Public Function FolderExists(folderName As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(folderName) And vbDirectory) = vbDirectory)
End Function

Public Function CombinePaths(p1 As String, p2 As String) As String
    CombinePaths = _
        TrimTrailingChars(p1, "/\") & "\" & _
        TrimLeadingChars(p2, "/\")
End Function

Public Function NormalizePath(ByVal p As String)
    Dim isUNC As Boolean
    isUNC = StartsWith(p, "\\")
    p = Replace(p, "/", "\")
    While InStr(p, "\\") > 0
        p = Replace(p, "\\", "\")
    Wend
    If isUNC Then p = "\" & p
    NormalizePath = TrimTrailingChars(p, "\")
End Function

Public Function GetDirectoryName(ByVal p As String)
    p = NormalizePath(p)
    Dim i As Integer
    i = InStrRev(p, "\")
    If i = 0 Then
        GetDirectoryName = ""
    Else
        GetDirectoryName = Left(p, i - 1)
    End If
End Function

Public Function GetFilename(ByVal p As String)
    p = NormalizePath(p)
    Dim i As Integer
    i = InStrRev(p, "\")
    GetFilename = Mid(p, i + 1)
End Function
