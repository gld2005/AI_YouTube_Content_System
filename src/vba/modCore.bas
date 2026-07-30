Attribute VB_Name = "modCore"
Option Explicit

Public Const PROJECT_TABLE_NAME As String = "VideoProjectsTable"
Public Const SCRIPT_INDEX_TABLE_NAME As String = "ScriptIndexTable"

Public Function GetTable(ByVal tableName As String) As ListObject
    Dim sheet As Worksheet
    For Each sheet In ThisWorkbook.Worksheets
        On Error Resume Next
        Set GetTable = sheet.ListObjects(tableName)
        On Error GoTo 0
        If Not GetTable Is Nothing Then Exit Function
    Next sheet
    Err.Raise vbObjectError + 1000, "GetTable", "Required table was not found: " & tableName
End Function

Public Function TableColumnIndex(ByVal table As ListObject, ByVal headerName As String) As Long
    Dim index As Long
    For index = 1 To table.ListColumns.Count
        If CStr(table.ListColumns(index).Name) = headerName Then
            TableColumnIndex = index
            Exit Function
        End If
    Next index
    Err.Raise vbObjectError + 1001, "TableColumnIndex", "Required table column was not found: " & headerName
End Function

Public Function FindTableRow(ByVal table As ListObject, ByVal idHeader As String, ByVal idValue As String) As ListRow
    Dim row As ListRow
    Dim columnIndex As Long
    columnIndex = TableColumnIndex(table, idHeader)
    For Each row In table.ListRows
        If StrComp(Trim$(CStr(row.Range.Cells(1, columnIndex).Value)), Trim$(idValue), vbTextCompare) = 0 Then
            Set FindTableRow = row
            Exit Function
        End If
    Next row
End Function

Public Function SheetExists(ByVal sheetName As String) As Boolean
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not sheet Is Nothing
    On Error GoTo 0
End Function

Public Function SafeSheetName(ByVal proposedName As String) As String
    Dim invalidCharacters As Variant
    Dim item As Variant
    Dim candidate As String
    invalidCharacters = Array("\", "/", ":", "*", "?", "[", "]")
    candidate = Trim$(proposedName)
    For Each item In invalidCharacters
        candidate = Replace(candidate, CStr(item), "-")
    Next item
    If Len(candidate) = 0 Then candidate = "Untitled"
    If Len(candidate) > 31 Then candidate = Left$(candidate, 31)
    SafeSheetName = candidate
End Function

Public Sub SetRowValue(ByVal row As ListRow, ByVal table As ListObject, ByVal headerName As String, ByVal value As Variant)
    row.Range.Cells(1, TableColumnIndex(table, headerName)).Value = value
End Sub

Public Sub ShowUserError(ByVal userMessage As String, ByVal technicalMessage As String)
    MsgBox userMessage, vbExclamation, "AI YouTube Content System"
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & technicalMessage
End Sub
