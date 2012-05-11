Attribute VB_Name = "ExcelUtils"
Option Explicit

Public Function IsWorkbookOpen(wbFilename As String) As Boolean
    Dim w As Workbook
    
    On Error GoTo notOpen
    Set w = Workbooks(wbFilename)
    IsWorkbookOpen = True
    Exit Function
    
notOpen:
    IsWorkbookOpen = False
End Function

Public Function SheetExists(sheetName As String, Optional wb As Workbook) As Boolean
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim s As Worksheet
    
    On Error GoTo notFound
    Set s = wb.Sheets(sheetName)
    SheetExists = True
    Exit Function
    
notFound:
    SheetExists = False
End Function

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

Public Function ExcelColNum(c As String) As Integer
    ExcelColNum = 0
    Dim i As Integer
    For i = 1 To Len(c)
        ExcelColNum = (ExcelColNum + Asc(Mid(c, i, 1)) - 64)
        If i < Len(c) Then ExcelColNum = ExcelColNum * 26
    Next
End Function

Function CellReference(ByVal r As Long, ByVal c As Integer, Optional sheet As String = "", _
    Optional absoluteRow As Boolean = False, Optional absoluteCol As Boolean = False) As String
    
    Dim ref As String
    ref = IIf(absoluteCol, "$", "") & ExcelCol(c) & IIf(absoluteRow, "$", "") & r
    If sheet = "" Then
        CellReference = ref
    Else
        CellReference = "'" & sheet & "'!" & ref
    End If
End Function

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

Public Sub ShowStatusMessage(statusMessage As String)
    Application.StatusBar = statusMessage
    Application.Caption = Len(statusMessage) & ":" & statusMessage
End Sub
