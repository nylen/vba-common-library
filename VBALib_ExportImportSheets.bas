Attribute VB_Name = "VBALib_ExportImportSheets"
' Common VBA Library - ExportImportSheets
' Provides functions for moving sheets of one workbook to another workbook,
' breaking links along the way.  Useful for sending the results of a workbook
' to others, or for saving key data items to another workbook.

Option Explicit

Const COL_FOLDER = 1
Const COL_FILENAME = 2
Const COL_SHEETNAME = 3
Const COL_NEWSHEETNAME = 4

' Imports Excel sheets from one or more workbooks into the given workbook,
' breaking any links and converting them to values.
' @param sheetsSpec: A two-dimensional array that describes the sheets to
' import into the workbook.  It should have one or more indices in its first
' dimension, and indices 1-3 (or 1-4) in its second dimension as follows:
'  - Index 1 (Folder) is the folder that this row's workbook appears in.
'  - Index 2 (Filename) is the filename of this row's workbook.
'  - Index 3 (SheetName) is the sheet to extract from this row's workbook.
'  - Index 4 (NewSheetName) is optional.  If present, the sheet described
'    by this row will be renamed to the NewSheetName after extraction.
' @param wb: The workbook that will receive the imported sheets (defaults to
' the workbook that contains this code).
Public Sub ImportExcelSheets(sheetsSpec() As Variant, Optional wb As Workbook)
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    CopyExcelSheets wb, sheetsSpec, False
End Sub

' Exports Excel sheets from one or more workbooks into a new workbook,
' breaking any links and converting them to values.  Returns True if the
' workbook was saved, or False if it was not saved.
' @param sheetsSpec: A two-dimensional array that describes the sheets to
' export to the new workbook.  This array has the same structure as in
' ImportExcelSheets, with the added feature that if the folder is blank and
' the filename is the string ThisWorkbook then the sheet to be exported will
' come from the current workbook.
' @param wb: The workbook that will receive the imported sheets (defaults to
' the active workbook).
Public Function ExportExcelSheets(sheetsSpec() As Variant, _
    wbFilename As String, Optional oAction As OverwriteAction = oaPrompt, _
    Optional openReadOnly As Boolean = False) As Boolean
    
    Dim wb As Workbook
    Set wb = Workbooks.Add
    
    While wb.Sheets.Count > 1
        DeleteSheet wb.Sheets(2)
    Wend
    
    CopyExcelSheets wb, sheetsSpec, True
    
    DeleteSheet wb.Sheets(1)
    
    ExportExcelSheets = SaveWorkbookAs(wb, wbFilename, oAction, openReadOnly)
End Function

Private Sub CopyExcelSheets(wb As Workbook, sheetsSpec() As Variant, _
    allowThisWorkbook As Boolean)
    
    Dim prevActiveSheet As Worksheet
    Set prevActiveSheet = wb.ActiveSheet
    
    Dim i1 As Long, i2 As Long
    i1 = LBound(sheetsSpec, 1)
    i2 = UBound(sheetsSpec, 1)
    
    ' The workbooks which contain the sheets we're interested in
    Dim wbFilenames() As String
    ReDim wbFilenames(i1 To i2)
    
    ' The names of the sheets we're interested in
    Dim sheetNames() As String
    ReDim sheetNames(i1 To i2)
    
    ' The new names of the sheets we're interested in
    Dim newSheetNames() As String
    ReDim newSheetNames(i1 To i2)
    
    ' The desired position of each sheet (this array stores a sheet name that
    ' the sheet will be placed immediately after, or the empty string if the
    ' sheet should be placed at the beginning of the workbook)
    Dim sheetPositions() As String
    ReDim sheetPositions(i1 To i2)
    
    Dim i As Long
    For i = i1 To i2
        Dim thisFolderName As String, thisFilename As String
        thisFolderName = CStr(sheetsSpec(i, COL_FOLDER))
        thisFilename = CStr(sheetsSpec(i, COL_FILENAME))
        If thisFolderName = "" Or thisFilename = "" Then
            wbFilenames(i) = thisFolderName & thisFilename
        Else
            wbFilenames(i) = NormalizePath(CombinePaths( _
                thisFolderName, thisFilename))
        End If
        
        If LCase(wbFilenames(i)) = "thisworkbook" Then
            If allowThisWorkbook Then
                wbFilenames(i) = "ThisWorkbook"
            Else
                Err.Raise 32000, Description:= _
                    "Cannot import Excel sheets from ThisWorkbook."
            End If
        End If
        
        sheetNames(i) = sheetsSpec(i, COL_SHEETNAME)
        
        If COL_NEWSHEETNAME <= UBound(sheetsSpec, 2) Then
            newSheetNames(i) = sheetsSpec(i, COL_NEWSHEETNAME)
        Else
            newSheetNames(i) = ""
        End If
        If newSheetNames(i) = "" Then
            newSheetNames(i) = sheetNames(i)
        End If
        
        If SheetExists(newSheetNames(i), wb) Then
            Dim sheetIndex As Long
            sheetIndex = wb.Sheets(newSheetNames(i)).Index
            If sheetIndex = 1 Then
                sheetPositions(i) = ""
            Else
                sheetPositions(i) = wb.Sheets(sheetIndex - 1).Name
            End If
        ElseIf i = i1 Then
            sheetPositions(i) = wb.Sheets(wb.Sheets.Count).Name
        Else
            sheetPositions(i) = newSheetNames(i - 1)
        End If
    Next
    
    Dim currentFilename As String
    Dim currentWb As Workbook
    Dim filesProcessed As New VBALib_List
    Dim sheetsToCopy As New VBALib_List
    
    Do
        currentFilename = ""
        sheetsToCopy.Clear
        
        For i = i1 To i2
            If currentFilename = "" Then
                If Not filesProcessed.Contains(LCase(wbFilenames(i))) Then
                    currentFilename = wbFilenames(i)
                    filesProcessed.Add LCase(wbFilenames(i))
                    Set currentWb = Nothing
                End If
            End If
            
            If LCase(currentFilename) = LCase(wbFilenames(i)) Then
                If currentWb Is Nothing Then
                    If currentFilename = "ThisWorkbook" Then
                        Set currentWb = ThisWorkbook
                    Else
                        ShowStatusMessage "Opening workbook:  " _
                            & GetFilename(currentFilename)
                        If IsWorkbookOpen(GetFilename(currentFilename)) Then
                            Set currentWb = Workbooks( _
                                GetFilename(currentFilename))
                        Else
                            Set currentWb = Workbooks.Open( _
                                wbFilenames(i), ReadOnly:=True)
                        End If
                        ClearStatusMessage
                    End If
                End If
                
                If SheetExists(newSheetNames(i), wb) Then
                    ShowStatusMessage "Deleting sheet:  " & newSheetNames(i)
                    DeleteSheetByName newSheetNames(i), wb
                    ClearStatusMessage
                End If
                
                ' Instead of copying sheets one at a time, save a list of the
                ' sheet names we need to copy and do them all at once.  This
                ' way is much faster.
                sheetsToCopy.Add sheetNames(i)
            End If
        Next
        
        If Not currentWb Is Nothing Then
            Dim oldSheetCount As Long
            oldSheetCount = wb.Sheets.Count
            
            ShowStatusMessage "Copying sheets from workbook:  " _
                & currentWb.Name
            currentWb.Sheets(sheetsToCopy.Items).Copy _
                After:=wb.Sheets(wb.Sheets.Count)
            
            ' Unhide any sheets that were hidden when copied over.
            For i = oldSheetCount + 1 To wb.Sheets.Count
                wb.Sheets(i).Visible = xlSheetVisible
            Next
            
            For i = i1 To i2
                If sheetNames(i) <> newSheetNames(i) Then
                    ShowStatusMessage "Renaming sheet:  " & newSheetNames(i)
                    wb.Sheets(sheetNames(i)).Name = newSheetNames(i)
                End If
            Next
            
            If ExcelLinkExists(currentFilename, wb) Then
                ShowStatusMessage "Breaking link to workbook:  " & currentWb.Name
                Dim currentWbLink As VBALib_ExcelLink
                Set currentWbLink = GetExcelLink(currentFilename, wb)
                currentWbLink.Break
            End If
            
            If currentFilename <> "ThisWorkbook" Then
                ShowStatusMessage "Closing workbook:  " & currentWb.Name
                currentWb.Close SaveChanges:=False
            End If
            Set currentWb = Nothing
            
            ClearStatusMessage
        End If
    Loop While currentFilename <> ""
    
    ShowStatusMessage "Rearranging sheets"
    For i = i1 To i2
        If sheetPositions(i) = "" Then
            wb.Sheets(newSheetNames(i)).Move _
                Before:=wb.Sheets(1)
        Else
            wb.Sheets(newSheetNames(i)).Move _
                After:=wb.Sheets(sheetPositions(i))
        End If
    Next
    ClearStatusMessage
    
    prevActiveSheet.Activate
End Sub
