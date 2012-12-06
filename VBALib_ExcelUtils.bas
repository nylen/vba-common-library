Attribute VB_Name = "VBALib_ExcelUtils"
' Common VBA Library, version 2012-10-03.1
' ExcelUtils
' Provides useful functions for working with the Excel object model.

Option Explicit

Private Declare Function CallNamedPipe Lib "kernel32" _
    Alias "CallNamedPipeA" ( _
        ByVal lpNamedPipeName As String, _
        ByVal lpInBuffer As Any, ByVal nInBufferSize As Long, _
        ByRef lpOutBuffer As Any, ByVal nOutBufferSize As Long, _
        ByRef lpBytesRead As Long, ByVal nTimeOut As Long) As Long

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Public Enum Corner
    cnrTopLeft
    cnrTopRight
    cnrBottomLeft
    cnrBottomRight
End Enum

' Determines whether a given workbook has been opened.  Pass this function
' a filename only, not a full path.
Public Function IsWorkbookOpen(wbFilename As String) As Boolean
    Dim w As Workbook
    
    On Error GoTo notOpen
    Set w = Workbooks(wbFilename)
    IsWorkbookOpen = True
    Exit Function
    
notOpen:
    IsWorkbookOpen = False
End Function

' Determines whether a sheet with a given name exists.
' @param wb: The workbook to check for the given sheet name (defaults to the
' active workbook).
Public Function SheetExists(sheetName As String, Optional wb As Workbook) _
    As Boolean
    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim s As Worksheet
    
    On Error GoTo notFound
    Set s = wb.Sheets(sheetName)
    SheetExists = True
    Exit Function
    
notFound:
    SheetExists = False
End Function

' Deletes the sheet with the given name, without prompting for confirmation.
' @param wb: The workbook to check for the given sheet name (defaults to the
' active workbook).
Public Sub DeleteSheetByName(sheetName As String, Optional wb As Workbook)
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If SheetExists(sheetName, wb) Then DeleteSheet wb.Sheets(sheetName)
End Sub

' Deletes the given worksheet, without prompting for confirmation.
Public Sub DeleteSheet(s As Worksheet)
    DeleteSheetOrSheets s
End Sub

' Deletes all sheets in the given Sheets object, without prompting for
' confirmation.
Public Sub DeleteSheets(s As Sheets)
    DeleteSheetOrSheets s
End Sub

Private Sub DeleteSheetOrSheets(s As Object)
    Dim prevDisplayAlerts As Boolean
    prevDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    s.Delete
    On Error GoTo 0
    Application.DisplayAlerts = prevDisplayAlerts
End Sub

' Returns the actual used range from a sheet.
' @param fromTopLeft: If True, returns the used range starting from cell A1,
' which is different from the way Excel's UsedRange property behaves if the
' workbook does not use any cells in the first row or column.
Public Function GetRealUsedRange(s As Worksheet, _
    Optional fromTopLeft As Boolean = True) As Range
    
    If fromTopLeft Then
        Set GetRealUsedRange = s.Range( _
            s.Cells(1, 1), _
            s.Cells( _
                s.UsedRange.Rows.Count + s.UsedRange.Row - 1, _
                s.UsedRange.Columns.Count + s.UsedRange.Column - 1))
    Else
        Set GetRealUsedRange = s.UsedRange
    End If
End Function

' Converts an integer column number to an Excel column string.
Public Function ExcelCol(c As Integer) As String
    ExcelCol = ExcelCol_ZeroBased(c - 1)
End Function

Private Function ExcelCol_ZeroBased(c As Integer) As String
    Dim c2 As Integer
    c2 = c \ 26
    If c2 = 0 Then
        ExcelCol_ZeroBased = Chr(65 + c)
    Else
        ExcelCol_ZeroBased = ExcelCol(c2) & Chr(65 + (c Mod 26))
    End If
End Function

' Converts an Excel column string to an integer column number.
Public Function ExcelColNum(c As String) As Integer
    ExcelColNum = 0
    Dim i As Integer
    For i = 1 To Len(c)
        ExcelColNum = (ExcelColNum + Asc(Mid(c, i, 1)) - 64)
        If i < Len(c) Then ExcelColNum = ExcelColNum * 26
    Next
End Function

' Builds an Excel cell reference.
Public Function CellReference(ByVal r As Long, ByVal c As Integer, _
    Optional sheet As String = "", Optional absoluteRow As Boolean = False, _
    Optional absoluteCol As Boolean = False) As String
    
    Dim ref As String
    ref = IIf(absoluteCol, "$", "") & ExcelCol(c) _
        & IIf(absoluteRow, "$", "") & r
    
    If sheet = "" Then
        CellReference = ref
    Else
        CellReference = "'" & sheet & "'!" & ref
    End If
End Function

' Returns a string describing the type of an Excel error value
' ("#DIV/0!", "#N/A", etc.)
Public Function ExcelErrorType(e As Variant) As String
    If IsError(e) Then
        Select Case e
            Case CVErr(xlErrDiv0)
                ExcelErrorType = "#DIV/0!"
            Case CVErr(xlErrNA)
                ExcelErrorType = "#N/A"
            Case CVErr(xlErrName)
                ExcelErrorType = "#NAME?"
            Case CVErr(xlErrNull)
                ExcelErrorType = "#NULL!"
            Case CVErr(xlErrNum)
                ExcelErrorType = "#NUM!"
            Case CVErr(xlErrRef)
                ExcelErrorType = "#REF!"
            Case CVErr(xlErrValue)
                ExcelErrorType = "#VALUE!"
            Case Else
                ExcelErrorType = "#UNKNOWN_ERROR"
        End Select
    Else
        ExcelErrorType = "(not an error)"
    End If
End Function

' Shows a status message to update the user on the progress of a long-running
' operation, in a way that can be detected by external applications.
Public Sub ShowStatusMessage(statusMessage As String)
    ' Show the message in the status bar.
    Application.StatusBar = statusMessage
    
    ' Set the Excel window title to the updated status message.  The window
    ' title as seen by the Windows API will then be:
    ' "Status Message - WorkbookFilename.xlsm"
    ' To allow external applications to extract just the status message,
    ' put the length of the message at the beginning.
    Application.Caption = Len(statusMessage) & ":" & statusMessage
End Sub

' Clears any status message that is currently being displayed by a macro.
Public Sub ClearStatusMessage()
    Application.StatusBar = False
    Application.Caption = Empty
End Sub

' Attempts to send a message to an external program that is running this macro
' and listening for messages.
Public Sub SendMessageToListener(msg As String)
    Dim bArray(0 To 0) As Byte
    Dim bytesRead As Long
    CallNamedPipe _
        "\\.\pipe\ExcelMacroCommunicationListener." & GetCurrentProcessId, _
        msg, Len(msg), bArray(0), 1, bytesRead, 500
End Sub

' Returns the cell in the given corner of the given range.
Public Function GetCornerCell(r As Range, c As Corner) As Range
    Select Case c
        Case cnrTopLeft
            Set GetCornerCell = r.Cells(1, 1)
        Case cnrTopRight
            Set GetCornerCell = r.Cells(1, r.Columns.Count)
        Case cnrBottomLeft
            Set GetCornerCell = r.Cells(r.Rows.Count, 1)
        Case cnrBottomRight
            Set GetCornerCell = r.Cells(r.Rows.Count, r.Columns.Count)
    End Select
End Function
