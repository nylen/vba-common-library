Attribute VB_Name = "VBAUtils"
Option Explicit

Public Function ModuleExists(moduleName As String, Optional wb As Workbook) As Boolean
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim c As Variant ' VBComponent
    
    On Error GoTo notFound
    Set c = wb.VBProject.VBComponents.Item(moduleName)
    ModuleExists = True
    Exit Function
    
notFound:
    ModuleExists = False
End Function

Public Sub RemoveModule(moduleName As String, Optional wb As Workbook)
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If Not ModuleExists(moduleName, wb) Then
        Err.Raise 32000, _
            Description:="Module '" & moduleName & "' not found."
    End If
    Dim c As Variant ' VBComponent
    Set c = wb.VBProject.VBComponents.Item(moduleName)
    wb.VBProject.VBComponents.Remove c
End Sub

Public Sub ExportModule(moduleName As String, filename As String, Optional wb As Workbook)
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If Not ModuleExists(moduleName, wb) Then
        Err.Raise 32000, _
            Description:="Module '" & moduleName & "' not found."
    End If
    wb.VBProject.VBComponents.Item(moduleName).Export filename
End Sub

Public Sub ImportModule(filename As String, Optional wb As Workbook)
    If wb Is Nothing Then Set wb = ActiveWorkbook
    wb.VBProject.VBComponents.Import filename
End Sub
