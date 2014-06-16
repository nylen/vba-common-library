VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBALib_ExcelTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Common VBA Library - VBALib_ExcelTable
' This is an object that represents and manipulates a table in an Excel
' workbook.

Option Explicit

Private mListObject As ListObject

' Returns the name of this Excel table.
Public Property Get Name() As String
    Name = mListObject.Name
End Property

' Returns the number of data rows in this Excel table, not including the
' header, total, or Insert rows.
Public Property Get DataRowCount() As Long
    DataRowCount = mListObject.ListRows.Count
End Property

' Returns the underlying object for this Excel table (a ListObject).
Public Property Get ListObject() As ListObject
    Set ListObject = mListObject
End Property

' Returns the value of the given cell in this Excel table.
Public Property Get CellValue(r As Long, c As Variant) As Variant
    CellValue = ColumnRange(c).Cells(r).Value
End Property

' Sets the value of the given cell in this Excel table.
Public Property Let CellValue(r As Long, c As Variant, val As Variant)
    ColumnRange(c).Cells(r).Value = val
End Property

' Returns the ListColumn object for the given column of this Excel table.
Private Property Get Column(c As Variant) As ListColumn
    On Error GoTo badColumn
    Set Column = mListObject.ListColumns(c)
    Exit Property
    
badColumn:
    On Error GoTo 0
    Err.Raise 32000, Description:= _
        "Table '" & Name & "' does not contain a column '" & c & "'."
End Property

' Returns the data range for the given column of this Excel table.
Public Property Get ColumnRange(c As Variant) As Range
    Dim listCol As ListColumn
    Set listCol = Column(c)
    
    On Error GoTo noDataRange
    Set ColumnRange = listCol.DataBodyRange
    Exit Property
    
noDataRange:
    On Error GoTo 0
    Err.Raise 32000, Description:= _
        "Failed to get data range for column '" & c & "' of table '" _
            & Name & "'."
End Property

' Initializes the table object with the necessary parameters.  INTERNAL ONLY -
' Do not call this method from user code.
Public Sub Initialize(listObj As ListObject)
    Set mListObject = listObj
End Sub

' Resizes this Excel table to the given number of data rows.
Public Sub Resize(numRows As Long)
    Dim oldNumRows As Long
    oldNumRows = DataRowCount
    
    ' Don't allow tables to be resized to zero rows.  Excel won't really do
    ' this anyway - it's possible to get the DataBodyRange to be empty, in
    ' which case the single row displayed in the table is the Insert row, but
    ' this creates more problems than it solves.  Instead, resize the table to
    ' one row, and set a flag to clear out any remaining data later.
    Dim clearTable As Boolean
    If numRows = 0 Then
        numRows = 1
        clearTable = True
    Else
        clearTable = False
    End If
    
    ' Resize the table (add 1 to the number of rows because mListObject.Range
    ' includes the header row).
    mListObject.Resize _
        mListObject.Range.Resize( _
            numRows + 1, _
            mListObject.ListColumns.Count)
    
    ' If the table is resized to have one row, but the row contains no data,
    ' the row will be treated as the Insert row, and the data row count will
    ' remain zero.  This will cause problems since the table doesn't actually
    ' have a DataBodyRange.  To avoid this situation, put a space in the first
    ' column, which will cause the Insert row to change to a data row.  After
    ' setting the value once, it can be removed and the row will remain part
    ' of the DataBodyRange.
    If numRows = 1 And DataRowCount = 0 Then
        mListObject.Range.Cells(1, 1).Offset(1, 0).Value = " "
        mListObject.DataBodyRange.ClearContents
    End If
    
    ' The user requested that the table be resized to zero rows.  We resized
    ' it down to one row, now clear out the data.
    If clearTable Then
        mListObject.DataBodyRange.ClearContents
        ClearSort
    End If
    
    ' If the new number of rows is less than the old number of rows, clear out
    ' the rows that were just removed from the table.
    If numRows < oldNumRows Then
        mListObject.DataBodyRange _
            .Offset(numRows, 0) _
            .Resize(oldNumRows - numRows, mListObject.ListColumns.Count) _
            .ClearContents
    End If
End Sub

' Clears this Excel table and resizes it to one row.
Public Sub Clear()
    Resize 0
End Sub

' Returns the values of this table as an array, optionally limited to a subset
' of the table's columns.
' @param startColumn: The index or name of the first column in the table to
' return the values for (defaults to 1).
' @param endColumn: The index or name of the last column in the table to
' return the values for (defaults to the last column in the table).
Public Function GetValues( _
    Optional startColumn As Variant = 1, _
    Optional endColumn As Variant = Empty) _
    As Variant()
    
    startColumn = Column(startColumn).Index
    If IsEmpty(endColumn) Then
        endColumn = mListObject.ListColumns.Count
    Else
        endColumn = Column(endColumn).Index
    End If
    
    GetValues = mListObject.DataBodyRange _
        .Resize(ColumnSize:=endColumn - startColumn + 1) _
        .Offset(ColumnOffset:=startColumn - 1) _
        .Value
End Function

' Sets the values of this table from an array, optionally resizing the table
' if the number of rows in the array does not match the number of rows in the
' table.
' @param resizeTable: Whether to resize the table to match the number of
' elements in the first dimension of the given array (defaults to True).
' @param startColumn: The index or name of the first column in the table that
' should receive values from the given array (defaults to 1).
Public Sub SetValues(arr() As Variant, _
    Optional resizeTable As Boolean = True, _
    Optional startColumn As Variant = 1)
    
    If resizeTable Then
        Resize ArrayLen(arr)
    End If
    
    ClearSort
    
    Dim startColumnNum As Long
    startColumnNum = Column(startColumn).Index
    
    Dim numRows As Long, numColumns As Long
    numRows = Min(ArrayLen(arr, 1), DataRowCount)
    If ArrayRank(arr) = 1 Then
        numColumns = 1
    Else
        numColumns = Min( _
            ArrayLen(arr, 2), _
            mListObject.ListColumns.Count - startColumnNum + 1)
    End If
    
    mListObject.DataBodyRange _
        .Offset(0, startColumnNum - 1) _
        .Resize(numRows, numColumns) _
        .Value = arr
End Sub

' Sorts this Excel table by the given column(s).
' @param columnSpecs: An array of or more column indices to sort by.  These
' can be numbers (1-based column indices) or strings (column names).  Column
' names can have ":asc" or ":desc" appended to them to sort in ascending or
' descending order (the default is ascending).
Public Sub Sort(columnSpecs As Variant)
    If Not IsArray(columnSpecs) Then
        Err.Raise 32000, Description:= _
            "When sorting a table, the column specifiers must be an array."
    End If
    
    columnSpecs = NormalizeArray(columnSpecs)
    If ArrayLen(columnSpecs) = 0 Then
        Err.Raise 32000, Description:= _
            "When sorting a table, at least one sort field must be given."
    End If
    
    With mListObject.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .SortFields.Clear
        
        Dim i As Integer, columnIndex As Variant, sortOrder As XlSortOrder
        For i = 1 To UBound(columnSpecs)
            columnIndex = columnSpecs(i)
            If EndsWith(CStr(columnIndex), ":asc", False) Then
                columnIndex = Left(columnIndex, Len(columnIndex) - 4)
                sortOrder = xlAscending
            ElseIf EndsWith(CStr(columnIndex), ":desc", False) Then
                columnIndex = Left(columnIndex, Len(columnIndex) - 5)
                sortOrder = xlDescending
            Else
                sortOrder = xlAscending
            End If
            
            .SortFields.Add Key:=ColumnRange(columnIndex), _
                SortOn:=xlSortOnValues, Order:=sortOrder
            
        Next
        
        .Apply
    End With
End Sub

' Clears any sorting applied to this Excel table.
Public Sub ClearSort()
    mListObject.Sort.SortFields.Clear
End Sub

' Creates a regular named range for each of this Excel table's columns.  This
' is useful if other workbooks need to link to this table, since references to
' table formulas like TblData[ColName] don't work with closed workbooks.
' @param namePrefix: The prefix which will be prepended to the names of the
' named ranges created from this table.
Public Sub CreateNamedRanges(namePrefix As String)
    Dim c As Long
    For c = 1 To mListObject.ListColumns.Count
        ' Get sheet then workbook
        mListObject.Parent.Parent.Names.Add _
            Name:=namePrefix & Column(c).Name, _
            RefersTo:="='" & mListObject.Parent.Name & "'!" _
                & ColumnRange(c).Address(True, True)
    Next
End Sub
