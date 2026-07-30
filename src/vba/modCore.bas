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

Public Function FindTableRow(ByVal table As ListObject, ByVal idColumn As Long, ByVal idValue As String) As ListRow
    Dim row As ListRow
    For Each row In table.ListRows
        If StrComp(Trim$(CStr(row.Range.Cells(1, idColumn).Value)), Trim$(idValue), vbTextCompare) = 0 Then
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
    invalidCharacters = Array("\", "/", ":", "*", "?", "[", "]")
    For Each item In invalidCharacters
        proposedName = Replace(proposedName, CStr(item), "-")
    Next item
    If Len(Trim$(proposedName)) = 0 Then proposedName = "Untitled"
    SafeSheetName = Left$(Trim$(proposedName), 31)
End Function

Public Sub SetColumnValue(ByVal row As ListRow, ByVal columnIndex As Long, ByVal value As Variant)
    row.Range.Cells(1, columnIndex).Value = value
End Sub

Public Sub ShowUserError(ByVal userMessage As String, ByVal technicalMessage As String)
    MsgBox userMessage, vbExclamation, "AI YouTube Content System"
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & technicalMessage
End Sub
