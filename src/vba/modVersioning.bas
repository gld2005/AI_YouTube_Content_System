Attribute VB_Name = "modVersioning"
Option Explicit

Public Sub SaveNewVersion()
    Dim projects As ListObject
    Dim projectRow As ListRow
    Dim videoId As String
    Dim currentSheetName As String
    Dim previousSheetName As String
    Dim currentSheet As Worksheet
    Dim previousSheet As Worksheet
    Dim reason As String

    videoId = Trim$(InputBox("请输入要保存版本的 Video ID：", "保存脚本新版本"))
    If Len(videoId) = 0 Then Exit Sub
    Set projects = GetTable(PROJECT_TABLE_NAME)
    Set projectRow = FindTableRow(projects, "Video ID", videoId)
    If projectRow Is Nothing Then
        ShowUserError "找不到对应的视频项目。", "Version save failed because the project was not found: " & videoId
        Exit Sub
    End If
    currentSheetName = CStr(projectRow.Range.Cells(1, TableColumnIndex(projects, "当前脚本页")).Value)
    previousSheetName = CStr(projectRow.Range.Cells(1, TableColumnIndex(projects, "上一版脚本页")).Value)
    If Not SheetExists(currentSheetName) Or Not SheetExists(previousSheetName) Then
        ShowUserError "当前脚本页或上一版脚本页缺失。", "Version save failed because a script page is missing."
        Exit Sub
    End If
    reason = Trim$(InputBox("请输入本次保存原因（可留空）：", "保存脚本新版本"))
    If MsgBox("将以当前脚本覆盖上一版。此操作不会自动在打开或关闭文件时执行。是否继续？", vbYesNo + vbQuestion, "保存脚本新版本") <> vbYes Then Exit Sub

    Set currentSheet = ThisWorkbook.Worksheets(currentSheetName)
    Set previousSheet = ThisWorkbook.Worksheets(previousSheetName)
    previousSheet.Cells.Clear
    currentSheet.UsedRange.Copy Destination:=previousSheet.Range("A1")
    previousSheet.Name = previousSheetName
    SetRowValue projectRow, projects, "最后更新", Date
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn") & " | Version saved for " & videoId & IIf(Len(reason) > 0, " | " & reason, "")
    MsgBox "已保存上一版脚本。段落差异标记将在脚本段落模型启用后自动应用。", vbInformation, "保存脚本新版本"
End Sub
