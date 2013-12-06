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
    
    wb.Close SaveChanges:=False
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
    
    ' The order in which sheets need to be moved when they are rearranged.  To
    ' see why this is necessary, imagine that a workbook contains sheets A, B,
    ' C, and D, but the program obtains these sheets in the order A, C, B, D.
    ' When rearranging sheets, A would be moved to its position (correctly),
    ' then C would be moved to its position after B, but since B was not in the
    ' desired position, then C would not be moved to its desired position
    ' either.  To solve this, store the order of the existing sheets in the
    ' workbook, and move the new sheets in that order.
    Dim sheetMoveOrder() As Long
    ReDim sheetMoveOrder(i1 To i2)
    ' Supporting variables for sheetMoveOrder.
    Dim sheetIndex As Long, sheetMoveOrderIndex As Long
    
    ' The list of Excel links to other workbooks that could not be broken.
    Dim linksFailedToBreak As New VBALib_List
    
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
        
        sheetMoveOrder(i) = i1 - 1
    Next
    
    ' Determine the order in which we need to rearrange sheets.  Start by
    ' looping over all of the workbook's current sheets, and checking if they
    ' are sheets that will be replaced during this run.  If so, then rearrange
    ' them in that order.
    sheetMoveOrderIndex = i1
    For sheetIndex = 1 To wb.Sheets.Count
        i = ArrayIndexOf(newSheetNames, wb.Sheets(sheetIndex).Name)
        If i >= i1 Then
            sheetMoveOrder(sheetMoveOrderIndex) = i
            sheetMoveOrderIndex = sheetMoveOrderIndex + 1
        End If
    Next
    ' Now, add any sheets that will be added to the workbook, but do not exist
    ' yet.  Ensure that these sheets are arranged according to the order in
    ' the specification passed to this function.
    For i = i1 To i2
        If Not ArrayContains(sheetMoveOrder, i) Then
            sheetMoveOrder(sheetMoveOrderIndex) = i
            sheetMoveOrderIndex = sheetMoveOrderIndex + 1
        End If
    Next
    ' Sanity check
    'If sheetMoveOrderIndex <> i2 + 1 Then Stop
    
    Dim currentFilename As String
    Dim currentWb As Workbook
    Dim filesProcessed As New VBALib_List
    Dim sheetsToCopy As New VBALib_List
    Dim oldLinkNames As New VBALib_List
    Dim newLinkNames As New VBALib_List
    
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
                        ShowStatusMessage "Opening workbook: " _
                            & GetFilename(currentFilename)
                        If IsWorkbookOpen(GetFilename(currentFilename)) Then
                            Set currentWb = Workbooks( _
                                GetFilename(currentFilename))
                        Else
                            Set currentWb = Workbooks.Open(wbFilenames(i), _
                                ReadOnly:=True, _
                                UpdateLinks:=False)
                        End If
                        ClearStatusMessage
                    End If
                End If
                
                If SheetExists(newSheetNames(i), wb) Then
                    ShowStatusMessage "Deleting sheet: " & newSheetNames(i)
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
            
            ' Store the list of linked files in the workbook before copying
            ' sheets over, because copying a sheet can add more than one link.
            oldLinkNames.Clear
            On Error Resume Next ' wb.LinkSources returns Empty if no links
            oldLinkNames.AddRange wb.LinkSources(xlExcelLinks)
            On Error GoTo 0
            
            ShowStatusMessage "Copying sheets from workbook: " _
                & currentWb.Name
            currentWb.Sheets(sheetsToCopy.Items).Copy _
                After:=wb.Sheets(wb.Sheets.Count)
            
            ' Unhide any sheets that were hidden when copied over.
            For i = oldSheetCount + 1 To wb.Sheets.Count
                wb.Sheets(i).Visible = xlSheetVisible
            Next
            
            For i = i1 To i2
                If LCase(currentFilename) = LCase(wbFilenames(i)) _
                    And sheetNames(i) <> newSheetNames(i) Then
                    
                    ShowStatusMessage "Renaming sheet: " & newSheetNames(i)
                    wb.Sheets(sheetNames(i)).Name = newSheetNames(i)
                End If
            Next
            
            ' Get the list of links again, and remove any that didn't exist
            ' before, as well as any link to the workbook that contains the
            ' sheet(s) we're currently copying.
            newLinkNames.Clear
            newLinkNames.AddRange wb.LinkSources(xlExcelLinks)
            If ExcelLinkExists(currentWb.Name, wb) Then
                newLinkNames.AddOnce currentWb.FullName
            End If
            
            Dim linkName_ As Variant, linkName As String
            For Each linkName_ In newLinkNames.Items
                linkName = linkName_
                ' Always try to remove the link to the current workbook, even
                ' if it already existed.
                If LCase(GetFilename(linkName)) = LCase(currentWb.Name) _
                    Or Not oldLinkNames.Contains(linkName) Then
                    
                    ShowStatusMessage "Breaking link to workbook: " _
                        & GetFilename(linkName)
                    Dim currentWbLink As VBALib_ExcelLink
                    Set currentWbLink = GetExcelLink(linkName, wb)
                    If Not currentWbLink.Break(False) Then
                        linksFailedToBreak.Add GetFilename(linkName)
                    End If
                End If
            Next
            
            If currentFilename <> "ThisWorkbook" Then
                ShowStatusMessage "Closing workbook: " & currentWb.Name
                currentWb.Close SaveChanges:=False
            End If
            Set currentWb = Nothing
            
            ClearStatusMessage
        End If
    Loop While currentFilename <> ""
    
    ShowStatusMessage "Rearranging sheets"
    For i = i1 To i2
        If sheetPositions(sheetMoveOrder(i)) = "" Then
            wb.Sheets(newSheetNames(sheetMoveOrder(i))).Move _
                Before:=wb.Sheets(1)
        ElseIf SheetExists(sheetPositions(sheetMoveOrder(i)), wb) Then
            wb.Sheets(newSheetNames(sheetMoveOrder(i))).Move _
                After:=wb.Sheets(sheetPositions(sheetMoveOrder(i)))
        End If
    Next
    ClearStatusMessage
    
    If linksFailedToBreak.Count Then
        MsgBox Prompt:="Failed to break links to one or more workbooks:" _
                & vbLf & vbLf & Join(linksFailedToBreak.Items, vbLf), _
            Title:="Excel link failure", _
            Buttons:=vbOKOnly + vbExclamation
    End If
    
    prevActiveSheet.Activate
End Sub
