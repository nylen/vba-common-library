Attribute VB_Name = "VBALib_FileUtils"
' Common VBA Library - FileUtils
' Provides useful functions for working with filenames and paths.

Option Explicit

Private Declare Function GetTempPathA Lib "kernel32" _
    (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

' Determines whether a file with the given name exists.
' @param findFolders: If true, the function will return true if a folder
' with the given name exists.
Public Function FileExists(ByVal testFilename As String, _
    Optional findFolders As Boolean = False) As Boolean
    
    ' Include read-only files, hidden files, system files.
    Dim attrs As Long
    attrs = (vbReadOnly Or vbHidden Or vbSystem)
    
    If findFolders Then
        attrs = (attrs Or vbDirectory) ' Include folders as well.
    End If
    
    'If Dir() returns something, the file exists.
    FileExists = False
    On Error Resume Next
    FileExists = (Dir(TrimTrailingChars(testFilename, "/\"), attrs) <> "")
End Function

' Determines whether a folder with the given name exists.
Public Function FolderExists(folderName As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(folderName) And vbDirectory) = vbDirectory)
End Function

' Creates the given directory, including any missing parent folders.
Public Sub MkDirRecursive(folderName As String)
    MkDirRecursiveInternal folderName, folderName
End Sub

Private Sub MkDirRecursiveInternal(folderName As String, _
    originalFolderName As String)
    
    If folderName = "" Then
        ' Too many recursive calls to this function (GetDirectoryName will
        ' eventually return an empty string)
        Err.Raise 32000, _
            Description:="Failed to create folder: " & originalFolderName
    End If
    
    Dim parentFolderName As String
    parentFolderName = GetDirectoryName(folderName)
    If Not FolderExists(parentFolderName) Then
        MkDirRecursiveInternal parentFolderName, originalFolderName
    End If
    
    If Not FolderExists(folderName) Then
        MkDir folderName
    End If
End Sub

' Merges two path components into a single path.
Public Function CombinePaths(p1 As String, p2 As String) As String
    CombinePaths = _
        TrimTrailingChars(p1, "/\") & "\" & _
        TrimLeadingChars(p2, "/\")
End Function

' Fixes slashes within a path:
'  - Converts all forward slashes to backslashes
'  - Removes multiple consecutive slashes (except for UNC paths)
'  - Removes any trailing slashes
Public Function NormalizePath(ByVal p As String) As String
    Dim isUNC As Boolean
    isUNC = StartsWith(p, "\\")
    p = Replace(p, "/", "\")
    While InStr(p, "\\") > 0
        p = Replace(p, "\\", "\")
    Wend
    If isUNC Then p = "\" & p
    NormalizePath = TrimTrailingChars(p, "\")
End Function

' Returns the folder name of a path (removes the last component
' of the path).
Public Function GetDirectoryName(ByVal p As String) As String
    p = NormalizePath(p)
    Dim i As Integer
    i = InStrRev(p, "\")
    If i = 0 Then
        GetDirectoryName = ""
    Else
        GetDirectoryName = Left(p, i - 1)
    End If
End Function

' Returns the filename of a path (the last component of the path).
Public Function GetFilename(ByVal p As String) As String
    p = NormalizePath(p)
    Dim i As Integer
    i = InStrRev(p, "\")
    GetFilename = Mid(p, i + 1)
End Function

Private Function ListFiles_Internal(filePattern As String, attrs As Long) _
    As Variant()
    
    Dim filesList As New VBALib_List
    Dim folderName As String
    
    If FolderExists(filePattern) Then
        filePattern = NormalizePath(filePattern) & "\"
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
                
                filesList.Add folderName & currFilename
            End If
        Else
            filesList.Add folderName & currFilename
        End If
        currFilename = Dir
    Wend
    
    ListFiles_Internal = filesList.Items
End Function

' Lists all files matching the given pattern.
' @param filePattern: A directory name, or a path with wildcards:
'  - C:\Path\to\Folder
'  - C:\Path\to\Folder\ExcelFiles.xl*
Public Function ListFiles(filePattern As String) As Variant()
    ListFiles = ListFiles_Internal(filePattern, _
        vbReadOnly Or vbHidden Or vbSystem)
End Function

' Lists all folders matching the given pattern.
' @param folderPattern: A directory name, or a path with wildcards:
'  - C:\Path\to\Folder
'  - C:\Path\to\Folder\OtherFolder_*
Public Function ListFolders(folderPattern As String) As Variant()
    ListFolders = ListFiles_Internal(folderPattern, _
        vbReadOnly Or vbHidden Or vbSystem Or vbDirectory)
End Function

' Returns the path to a folder that can be used to store temporary
' files.
Public Function GetTempPath() As String
    Const MAX_PATH = 256
    
    Dim folderName As String
    Dim ret As Long
    
    folderName = String(MAX_PATH, 0)
    ret = GetTempPathA(MAX_PATH, folderName)
    
    If ret <> 0 Then
        GetTempPath = Left(folderName, InStr(folderName, Chr(0)) - 1)
    Else
        Err.Raise 32000, Description:= _
            "Error getting temporary folder."
    End If
End Function
