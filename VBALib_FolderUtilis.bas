Attribute VB_Name = "VB_FileFolder"
Option Explicit
'Authored 2015-2017 by Jeremy Dean Gerdes <jeremy.gerdes@navy.mil>
     'Public Domain in the United States of America,
     'any international rights are waived through the CC0 1.0 Universal public domain dedication <https://creativecommons.org/publicdomain/zero/1.0/legalcode>
     'http://www.copyright.gov/title17/
     'In accrordance with 17 U.S.C. § 105 This work is 'noncopyright' or in the 'public domain'
         'Subject matter of copyright: United States Government works
         'protection under this title is not available for
         'any work of the United States Government, but the United States
         'Government is not precluded from receiving and holding copyrights
         'transferred to it by assignment, bequest, or otherwise.
     'as defined by 17 U.S.C § 101
         '...
         'A “work of the United States Government” is a work prepared by an
         'officer or employee of the United States Government as part of that
         'person’s official duties.
         '...
Public Function BuildDir(strPath) As Boolean
    On Error Resume Next
    Dim fso As Object ' As Scripting.FileSystemObject
    Dim arryPaths As Variant
    Dim strBuiltPath As String, intDir As Integer, fRestore As Boolean: fRestore = False
    If Left(strPath, 2) = "\\" Then
        strPath = Right(strPath, Len(strPath) - 2)
        fRestore = True
    End If
    Set fso = CreateObject("Scripting.FileSystemObject") ' New Scripting.FileSystemObject
    arryPaths = Split(strPath, "\")
    'Restore Server file path
    If fRestore Then
        arryPaths(0) = "\\" & arryPaths(0)
    End If
    For intDir = 0 To UBound(arryPaths)
        strBuiltPath = strBuiltPath & arryPaths(intDir)
        If Not fso.FolderExists(strBuiltPath) Then
            fso.CreateFolder strBuiltPath
        End If
        strBuiltPath = strBuiltPath & "\"
    Next
    BuildDir = (Err.Number = 0) 'True if no errors
End Function

Public Function GetRelativePathViaParentAlternateRoot(ByVal strRootPath As String, ByVal strDestination As String, Optional ByRef intParentCount As Integer)
    If Left(strDestination, 3) = "..\" Then
        intParentCount = intParentCount + 1
        strRootPath = Left(strRootPath, InStrRev(strRootPath, "\") - 1)
        strDestination = Right(strDestination, Len(strDestination) - 3)
        GetRelativePathViaParentAlternateRoot = GetRelativePathViaParentAlternateRoot(strRootPath, strDestination, intParentCount)
    ElseIf Left(strDestination, 1) = "\" And Not (Left(strDestination, 2) = "\\") Then
        strDestination = Right(strDestination, Len(strDestination) - 1)
    ElseIf Right(strDestination, 1) = "\" Then
        strDestination = Left(strDestination, Len(strDestination) - 1)
    End If
    If intParentCount <> -1 Then
        GetRelativePathViaParentAlternateRoot = StripTrailingBackSlash(strRootPath) & "\" & strDestination
    End If
    intParentCount = -1
End Function

Public Function GetRelativePathViaParent(Optional ByVal strPath)
'Usage for up 2 dirs is GetRelativePathViaParent("..\..\Destination")
Dim strCurrentPath As String, strVal As String
Dim oThisApplication As Object:    Set oThisApplication = Application
Dim fIsServerPath As Boolean: fIsServerPath = False
Dim aryCurrentFolder As Variant, aryParentPath As Variant
    Select Case True
        Case InStrRev(oThisApplication.Name, "Excel") > 0
            strCurrentPath = oThisApplication.ThisWorkbook.Path
        Case InStrRev(oThisApplication.Name, "Access") > 0
            strCurrentPath = oThisApplication.CurrentProject.Path
    End Select
    If Left(strCurrentPath, 2) = "\\" Then
        strCurrentPath = Right(strCurrentPath, Len(strCurrentPath) - 2)
        fIsServerPath = True
    End If
    aryCurrentFolder = Split(strCurrentPath, "\")
    If IsMissing(strPath) Then
        strPath = vbNullString
    End If
    aryParentPath = Split(strPath, "..\")
    If fIsServerPath Then
        aryCurrentFolder(0) = "\\" & aryCurrentFolder(0)
    End If
    Dim intDir As Integer, intParentCount As Integer
    If UBound(aryParentPath) = -1 Then
        intParentCount = 0
    Else
        intParentCount = UBound(aryParentPath)
    End If
    For intDir = 0 To UBound(aryCurrentFolder) - intParentCount
        strVal = strVal & aryCurrentFolder(intDir) & "\"
    Next
    strVal = StripTrailingBackSlash(strVal)
    If IsArrayAllocated(aryParentPath) Then
        strVal = strVal & "\" & aryParentPath(UBound(aryParentPath))
    End If
    GetRelativePathViaParent = strVal
End Function

Public Function StripTrailingBackSlash(ByRef strPath As String)
        If Right(strPath, 1) = "\" Then
            StripTrailingBackSlash = Left(strPath, Len(strPath) - 1)
        Else
            StripTrailingBackSlash = strPath
        End If
End Function
