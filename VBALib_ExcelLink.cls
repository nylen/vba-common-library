VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBALib_ExcelLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Common VBA Library - VBALib_ExcelLink
' This is an object that represents and manipulates a link from one Excel
' workbook to another.

Option Explicit

Private mWorkbook As Workbook
Private mName As String

Public Property Get Name() As String
    Name = mName
End Property

' Returns the current status of this Excel link.  The possible values can be
' found here:  https://office.microsoft.com/en-us/excel-help/HV080555653.aspx
Public Property Get Status() As XlLinkStatus
    Status = mWorkbook.LinkInfo(mName, xlLinkInfoStatus)
End Property

' Initializes the link object with the necessary parameters.  INTERNAL ONLY -
' Do not call this method from user code.
Public Sub Initialize(wb As Workbook, linkName As String)
    mName = linkName
    Set mWorkbook = wb
End Sub

' Updates the values for this link from the source spreadsheet.
Public Sub UpdateValues()
    If Status <> xlLinkStatusSourceOpen Then
        mWorkbook.UpdateLink mName, xlLinkTypeExcelLinks
    End If
End Sub

' Breaks this link, replacing the formulas that linked to the source
' spreadsheet with their values.  Returns whether the operation succeeded
' (sometimes Excel prevents breaking links until a workbook is closed and
' re-opened).
Public Function Break() As Boolean
    mWorkbook.BreakLink mName, xlLinkTypeExcelLinks
    Break = Not ExcelLinkExists(mName, mWorkbook)
End Function

' Changes this link to point to another Excel workbook.
Public Sub ChangeSource(newWorkbookPath As String)
    mWorkbook.ChangeLink mName, newWorkbookPath, xlLinkTypeExcelLinks
    If ExcelLinkExists(newWorkbookPath) Then
        mName = GetExcelLink(newWorkbookPath).Name
    Else
        mName = newWorkbookPath
    End If
End Sub

' Opens the Excel workbook referenced by this link.
Public Function OpenSource() As Workbook
    mWorkbook.OpenLinks mName, Type:=xlLinkTypeExcelLinks
End Function
