Attribute VB_Name = "modNavigation"
Option Explicit

Public Sub GoToWorkspace()
    Sheet3.Activate
    Sheet3.Range("A1").Select
End Sub

Public Sub GoToDashboard()
    Sheet2.Activate
    Sheet2.Range("A1").Select
End Sub

Public Sub GoToScriptFromActiveProject()
    Dim projects As ListObject
    Dim projectRow As ListRow
    Dim scriptSheetName As String
    Set projects = GetTable(PROJECT_TABLE_NAME)
    Set projectRow = FindTableRow(projects, 1, CStr(ActiveCell.Value))
    If projectRow Is Nothing Then
        ShowUserError "Select a valid Video ID in the project table first.", "No project was found for the active value."
        Exit Sub
    End If
    scriptSheetName = CStr(projectRow.Range.Cells(1, 28).Value)
    If Len(scriptSheetName) = 0 Or Not SheetExists(scriptSheetName) Then
        ShowUserError "The selected project does not have a script sheet yet.", "The project script sheet is missing."
        Exit Sub
    End If
    ThisWorkbook.Worksheets(scriptSheetName).Activate
End Sub
