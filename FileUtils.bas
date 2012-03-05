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

Private Function ListFiles_Internal(filePattern As String, attrs As Long) As Variant
    Dim files As New List
    Dim folderName As String
    
    If FolderExists(filePattern) Then
        filePattern = TrimTrailingChars(filePattern, "\/") & "\"
        folderName = filePattern
    Else
        folderName = GetDirectoryName(filePattern) & "\"
    End If
    
    Dim currFilename As String
    currFilename = Dir(filePattern, attrs)
    
    While currFilename <> ""
        If (attrs And vbDirectory) = vbDirectory Then
            If FolderExists(folderName & currFilename) _
                And currFilename <> "." And currFilename <> ".." Then
                
                files.Add folderName & currFilename
            End If
        Else
            files.Add folderName & currFilename
        End If
        currFilename = Dir
    Wend
    
    If files.HasItems Then
        ListFiles_Internal = files.Items
    Else
        ListFiles_Internal = Empty
    End If
End Function

Public Function ListFiles(filePattern As String)
    ListFiles = ListFiles_Internal(filePattern, _
        vbReadOnly Or vbHidden Or vbSystem)
End Function

Public Function ListFolders(folderPattern As String)
    ListFolders = ListFiles_Internal(folderPattern, _
        vbReadOnly Or vbHidden Or vbSystem Or vbDirectory)
End Function
