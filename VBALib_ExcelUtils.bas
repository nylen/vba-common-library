Attribute VB_Name = "VBALib_ExcelUtils"
' Common VBA Library - ExcelUtils
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

Public Enum OverwriteAction
    oaPrompt = 1
    oaOverwrite = 2
    oaSkip = 3
    oaError = 4
    oaCreateDirectory = 8
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

' Determines whether a sheet with the given name exists.
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

' Determines whether a chart with the given name exists.
' @param chartName: The name of the chart to check for.
' @param sheetName: The name of the worksheet that contains the given chart
' (optional; the default is to search all worksheets).
' @param wb: The workbook to check for the given chart name (defaults to the
' active workbook.
Public Function ChartExists(chartName As String, _
    Optional sheetName As String = "", Optional wb As Workbook) As Boolean
    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    
    Dim s As Worksheet
    Dim c As ChartObject
    
    ChartExists = False
    
    If sheetName = "" Then
        For Each s In wb.Sheets
            If ChartExists(chartName, s.Name, wb) Then
                ChartExists = True
                Exit Function
            End If
        Next
    Else
        Set s = wb.Sheets(sheetName)
        On Error GoTo notFound
        Set c = s.ChartObjects(chartName)
        ChartExists = True
notFound:
    End If
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
' sheet does not use any cells in the top row(s) and/or leftmost column(s).
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

' Sets the value of the given range if it is different than the proposed value.
' Returns whether the value of the range was changed.
Public Function SetValueIfNeeded(rng As Range, val As Variant) As Boolean
    If rng.Value = val Then
        SetValueIfNeeded = False
    Else
        rng.Value = val
        SetValueIfNeeded = True
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
        CellReference = "'" & Replace(sheet, "'", "''") & "'!" & ref
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

' Shows a status message for 2-3 seconds then removes it.
Public Sub FlashStatusMessage(statusMessage As String)
    ShowStatusMessage statusMessage
    Application.OnTime Now + TimeValue("0:00:02"), "ClearStatusMessage"
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

' Returns an array of objects representing the other Excel workbooks that the
' given workbook links to.
' @param wb: The source workbook (defaults to the active workbook).
Public Function GetAllExcelLinks(Optional wb As Workbook) As Variant
    If wb Is Nothing Then Set wb = ActiveWorkbook
    
    Dim linkNames() As Variant
    linkNames = NormalizeArray(ActiveWorkbook.LinkSources(xlExcelLinks))
    
    If ArrayLen(linkNames) Then
        Dim linksArr() As VBALib_ExcelLink
        ReDim linksArr(1 To ArrayLen(linkNames))
        Dim i As Integer
        For i = 1 To UBound(linkNames)
            Set linksArr(i) = New VBALib_ExcelLink
            linksArr(i).Initialize wb, CStr(linkNames(i))
        Next
        GetAllExcelLinks = linksArr
    Else
        GetAllExcelLinks = Array()
        Exit Function
    End If
End Function

Private Function GetMatchingLinkName(linkFilename As String, _
    Optional wb As Workbook) As String
    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    
    Dim linkNames() As Variant
    linkNames = NormalizeArray(wb.LinkSources(xlExcelLinks))
    
    Dim i As Integer, matchingLinkName As String
    
    ' First look for a link with the exact full path given by linkFilename
    For i = 1 To UBound(linkNames)
        If LCase(linkNames(i)) = LCase(linkFilename) Then
            GetMatchingLinkName = linkNames(i)
            Exit Function
        End If
    Next
    
    ' Next look for a link with the same filename as linkFilename.  Do it in
    ' two steps because it is actually possible for Excel to link to two
    ' workbooks with the same name in different folders.  No one should ever
    ' do this, but we'll try to support retrieving such links anyway.
    For i = 1 To UBound(linkNames)
        If LCase(GetFilename(linkNames(i))) = _
            LCase(GetFilename(linkFilename)) Then
            
            GetMatchingLinkName = linkNames(i)
            Exit Function
        End If
    Next
    
    GetMatchingLinkName = ""
End Function

' Returns an object representing the link to the Excel workbook with the given
' filename.
' @param linkFilename: The path or filename of the linked Excel workbook.
' @param wb: The workbook that contains the link (defaults to the active
' workbook).
Public Function GetExcelLink(linkFilename As String, Optional wb As Workbook) _
    As VBALib_ExcelLink
    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    
    Dim matchingLinkName As String
    matchingLinkName = GetMatchingLinkName(linkFilename, wb)
    
    If matchingLinkName = "" Then
        Err.Raise 32000, Description:= _
            "No Excel link exists with the given name ('" & linkFilename _
                & "')."
    Else
        Set GetExcelLink = New VBALib_ExcelLink
        GetExcelLink.Initialize wb, matchingLinkName
    End If
End Function

' Returns whether an Excel link matching the given workbook filename exists.
' @param wb: The workbook that contains the link (defaults to the active
' workbook).
Public Function ExcelLinkExists(linkFilename As String, _
    Optional wb As Workbook) As Boolean
    
    ExcelLinkExists = (GetMatchingLinkName(linkFilename, wb) <> "")
End Function

' Refreshes all Access database connections in the given workbook.
' @param wb: The workbook to refresh (defaults to the active workbook).
Public Sub RefreshAccessConnections(Optional wb As Workbook)
    If wb Is Nothing Then Set wb = ActiveWorkbook
    
    Dim cn As WorkbookConnection
    
    On Error GoTo err_
    Application.Calculation = xlCalculationManual
    
    Dim numConnections As Integer, i As Integer
    
    For Each cn In wb.Connections
        If cn.Type = xlConnectionTypeOLEDB Then
            numConnections = numConnections + 1
        End If
    Next
    
    For Each cn In wb.Connections
        If cn.Type = xlConnectionTypeOLEDB Then
            i = i + 1
            ShowStatusMessage "Refreshing data connection '" _
                & cn.OLEDBConnection.CommandText _
                & "' (" & i & " of " & numConnections & ")"
            cn.OLEDBConnection.BackgroundQuery = False
            cn.Refresh
       End If
    Next
    
    GoTo done_
err_:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    
done_:
    ShowStatusMessage "Recalculating"
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    
    ClearStatusMessage
End Sub

' Returns a wrapper object for the table with the given name in the given
' workbook.
' @param wb: The workbook that contains the table (defaults to the active
' workbook).
Public Function GetExcelTable(tblName As String, Optional wb As Workbook) _
    As VBALib_ExcelTable
    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    
    On Error GoTo notFound
    
    Dim wbPrevActive As Workbook
    Set wbPrevActive = ActiveWorkbook
    wb.Activate
    
    ' We could just do Range(tblName).ListObject, but then this would allow
    ' getting a table by any of its cells or columns.  Instead, do some
    ' verification that the string we were passed is actually the name of
    ' a table.
    Dim tbl As ListObject
    Set tbl = Range(tblName).Parent.ListObjects(tblName)
    
    wbPrevActive.Activate
    
    Set GetExcelTable = New VBALib_ExcelTable
    GetExcelTable.Initialize tbl
    Exit Function
    
notFound:
    On Error GoTo 0
    Err.Raise 32000, Description:= _
        "Could not find table '" & tblName & "'."
End Function

' Returns the Excel workbook format for the given file extension.
Public Function GetWorkbookFileFormat(fileExtension As String) As XlFileFormat
    Select Case LCase(Replace(fileExtension, ".", ""))
        Case "xls"
            GetWorkbookFileFormat = xlExcel8
        Case "xla"
            GetWorkbookFileFormat = xlAddIn8
        Case "xlt"
            GetWorkbookFileFormat = xlTemplate8
        Case "csv"
            GetWorkbookFileFormat = xlCSV
        Case "txt"
            GetWorkbookFileFormat = xlCurrentPlatformText
        Case "xlsx"
            GetWorkbookFileFormat = xlOpenXMLWorkbook
        Case "xlsm"
            GetWorkbookFileFormat = xlOpenXMLWorkbookMacroEnabled
        Case "xlsb"
            GetWorkbookFileFormat = xlExcel12
        Case "xlam"
            GetWorkbookFileFormat = xlOpenXMLAddIn
        Case "xltx"
            GetWorkbookFileFormat = xlOpenXMLTemplate
        Case "xltm"
            GetWorkbookFileFormat = xlOpenXMLTemplateMacroEnabled
        Case Else
            Err.Raise 32000, Description:= _
                "Unrecognized Excel file extension: '" & fileExtension & "'"
    End Select
End Function

' Saves the given workbook as a different filename, with options for handling
' the case where the file already exists.  Returns True if the workbook was
' saved, or False if it was not saved.
' @param oAction: The action that will be taken if the given file exists.  This
' parameter also accepts the oaCreateDirectory flag, which means that the
' directory hierarchy of the requested filename will be created if it does not
' already exist.  If not given, defaults to oaPrompt.
Public Function SaveWorkbookAs(wb As Workbook, newFilename As String, _
    Optional oAction As OverwriteAction = oaPrompt, _
    Optional openReadOnly As Boolean = False) As Boolean
    
    If Not FolderExists(GetDirectoryName(newFilename)) Then
        If oAction And oaCreateDirectory Then
            MkDirRecursive GetDirectoryName(newFilename)
        Else
            Err.Raise 32000, Description:= _
                "The parent folder of the requested workbook filename " _
                    & "does not exist:" & vbLf & vbLf & newFilename
        End If
    End If
    
    If FileExists(newFilename) Then
        If oAction And oaOverwrite Then
            Kill newFilename
            ' Proceed to save the file
            
        ElseIf oAction And oaError Then
            Err.Raise 32000, Description:= _
                "The given filename already exists:" _
                    & vbLf & vbLf & newFilename
            
        ElseIf oAction And oaPrompt Then
            Dim r As VbMsgBoxResult
            r = MsgBox(Title:="Overwrite Excel file?", _
                Buttons:=vbYesNo + vbExclamation, _
                Prompt:="The following Excel file already exists:" _
                    & vbLf & vbLf & newFilename & vbLf & vbLf _
                    & "Do you want to overwrite it?")
            If r = vbYes Then
                Kill newFilename
                ' Proceed to save the file
            Else
                SaveWorkbookAs = False
                Exit Function
            End If
            
        ElseIf oAction And oaSkip Then
            SaveWorkbookAs = False
            Exit Function
            
        Else
            Err.Raise 32000, Description:= _
                "Bad overwrite action value passed to SaveWorkbookAs."
            
        End If
    End If
    
    ' wb.SaveCopyAs doesn't take all the fancy arguments that wb.SaveAs does,
    ' but it's the only way to save a copy of the current workbook.  This
    ' means, among other things, that it is not possible to save the workbook
    ' as a different format than the original workbook.  To work around this,
    ' call SaveCopyAs with a temporary filename first, then open the temporary
    ' file, then call SaveAs with the desired filename and options.
    
    Dim tmpFilename As String
    tmpFilename = CombinePaths(GetTempPath, Int(Rnd * 1000000) & "-" & wb.Name)
    wb.SaveCopyAs tmpFilename
    
    Dim wbTmp As Workbook
    Set wbTmp = Workbooks.Open(tmpFilename, UpdateLinks:=False, ReadOnly:=True)
    wbTmp.SaveAs filename:=newFilename, _
        FileFormat:=GetWorkbookFileFormat(GetFileExtension(newFilename)), _
        ReadOnlyRecommended:=openReadOnly
    wbTmp.Close SaveChanges:=False
    
    Kill tmpFilename
    
    SaveWorkbookAs = True
End Function

' Saves the current workbook as a different filename, with options for handling
' the case where the file already exists.  Returns True if the workbook was
' saved, or False if it was not saved.
' @param oAction: The action that will be taken if the given file exists.  This
' parameter also accepts the oaCreateDirectory flag, which means that the
' directory hierarchy of the requested filename will be created if it does not
' already exist.  If not given, defaults to oaPrompt.
' @param openReadOnly: True or False to determine whether the created workbook
' will prompt users to open it as read-only (defaults to False).
Public Function SaveActiveWorkbookAs(newFilename As String, _
    Optional oAction As OverwriteAction = oaPrompt, _
    Optional openReadOnly As Boolean = False) As Boolean
    
    SaveActiveWorkbookAs = SaveWorkbookAs(ActiveWorkbook, newFilename, _
        oAction, openReadOnly)
End Function
