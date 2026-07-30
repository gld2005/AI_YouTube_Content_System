Attribute VB_Name = "modNavigation"
Option Explicit

Public Sub GoToWorkspace()
    ThisWorkbook.Worksheets("内容工作台").Activate
    ThisWorkbook.Worksheets("内容工作台").Range("A1").Select
End Sub

Public Sub GoToDashboard()
    ThisWorkbook.Worksheets("总览").Activate
    ThisWorkbook.Worksheets("总览").Range("A1").Select
End Sub

Public Sub GoToScriptFromActiveProject()
    Dim projects As ListObject
    Dim row As ListRow
    Dim scriptSheetName As String
    Set projects = GetTable(PROJECT_TABLE_NAME)
    Set row = FindTableRow(projects, "Video ID", ActiveCell.Value)
    If row Is Nothing Then
        ShowUserError "请先在“视频项目”中选择有效的 Video ID。", "No project was found for the active value."
        Exit Sub
    End If
    scriptSheetName = CStr(row.Range.Cells(1, TableColumnIndex(projects, "当前脚本页")).Value)
    If Len(scriptSheetName) = 0 Or Not SheetExists(scriptSheetName) Then
        ShowUserError "该项目尚未创建脚本页。", "The project script sheet is missing."
        Exit Sub
    End If
    ThisWorkbook.Worksheets(scriptSheetName).Activate
End Sub
