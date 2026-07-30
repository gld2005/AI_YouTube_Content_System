Attribute VB_Name = "modNewVideo"
Option Explicit

Public Sub NewVideoScript()
    Dim videoId As String
    Dim videoTitle As String
    Dim projects As ListObject
    Dim scriptIndex As ListObject
    Dim projectRow As ListRow
    Dim scriptRow As ListRow
    Dim existingProject As ListRow
    Dim currentSheetName As String
    Dim previousSheetName As String

    videoId = Trim$(InputBox("请输入 Video ID，例如 YT-002：", "新建视频脚本"))
    If Len(videoId) = 0 Then Exit Sub
    videoTitle = Trim$(InputBox("请输入视频标题：", "新建视频脚本"))
    If Len(videoTitle) = 0 Then
        ShowUserError "视频标题不能为空。", "New video creation was cancelled because the title was empty."
        Exit Sub
    End If

    Set projects = GetTable(PROJECT_TABLE_NAME)
    Set projectRow = FindTableRow(projects, "Video ID", videoId)
    If Not projectRow Is Nothing Then
        ShowUserError "该 Video ID 已存在，未创建重复项目。", "Duplicate Video ID: " & videoId
        Exit Sub
    End If

    currentSheetName = SafeSheetName(videoId & "｜脚本")
    previousSheetName = SafeSheetName(videoId & "｜上一版")
    If SheetExists(currentSheetName) Or SheetExists(previousSheetName) Then
        ShowUserError "脚本页名称已存在，请使用其他 Video ID。", "Duplicate script sheet name for Video ID: " & videoId
        Exit Sub
    End If

    Set projectRow = projects.ListRows.Add
    SetRowValue projectRow, projects, "Video ID", videoId
    SetRowValue projectRow, projects, "视频标题", videoTitle
    SetRowValue projectRow, projects, "Series ID", "SER-UNCLASSIFIED"
    SetRowValue projectRow, projects, "项目状态", "策划中"
    SetRowValue projectRow, projects, "脚本状态", "未开始"
    SetRowValue projectRow, projects, "优先级", "P1"
    SetRowValue projectRow, projects, "开始日期", Date
    SetRowValue projectRow, projects, "最后更新", Date
    SetRowValue projectRow, projects, "当前脚本页", currentSheetName
    SetRowValue projectRow, projects, "上一版脚本页", previousSheetName
    SetRowValue projectRow, projects, "当前版本", "V1"
    SetRowValue projectRow, projects, "主项目", "否"

    CreateScriptSheets currentSheetName, previousSheetName

    Set scriptIndex = GetTable(SCRIPT_INDEX_TABLE_NAME)
    Set scriptRow = scriptIndex.ListRows.Add
    SetRowValue scriptRow, scriptIndex, "Video ID", videoId
    SetRowValue scriptRow, scriptIndex, "视频标题", videoTitle
    SetRowValue scriptRow, scriptIndex, "Series ID", "SER-UNCLASSIFIED"
    SetRowValue scriptRow, scriptIndex, "当前脚本页", currentSheetName
    SetRowValue scriptRow, scriptIndex, "上一版脚本页", previousSheetName
    SetRowValue scriptRow, scriptIndex, "当前版本", "V1"
    SetRowValue scriptRow, scriptIndex, "脚本状态", "未开始"
    SetRowValue scriptRow, scriptIndex, "最后更新", Date

    ThisWorkbook.Worksheets(currentSheetName).Activate
    MsgBox "已创建新视频项目和脚本页。请确认系列与继承选项后继续。", vbInformation, "新建视频脚本"
End Sub

Private Sub CreateScriptSheets(ByVal currentSheetName As String, ByVal previousSheetName As String)
    Dim templateSheet As Worksheet
    Set templateSheet = ThisWorkbook.Worksheets("脚本模板")
    templateSheet.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    ActiveSheet.Name = currentSheetName
    templateSheet.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    ActiveSheet.Name = previousSheetName
    ActiveSheet.Range("A1").Value = "脚本上一版｜待首次保存版本"
End Sub

Public Function CreateVideoForSmokeTest(ByVal videoId As String, ByVal videoTitle As String) As Boolean
    Dim projects As ListObject
    Dim scriptIndex As ListObject
    Dim projectRow As ListRow
    Dim scriptRow As ListRow
    Dim currentSheetName As String
    Dim previousSheetName As String

    On Error GoTo Failed
    Set projects = GetTable(PROJECT_TABLE_NAME)
    Set existingProject = FindTableRow(projects, "Video ID", videoId)
    If Not existingProject Is Nothing Then GoTo Failed
    currentSheetName = SafeSheetName(videoId & "｜脚本")
    previousSheetName = SafeSheetName(videoId & "｜上一版")
    If SheetExists(currentSheetName) Or SheetExists(previousSheetName) Then GoTo Failed

    Set projectRow = projects.ListRows.Add
    SetRowValue projectRow, projects, "Video ID", videoId
    SetRowValue projectRow, projects, "视频标题", videoTitle
    SetRowValue projectRow, projects, "Series ID", "SER-UNCLASSIFIED"
    SetRowValue projectRow, projects, "项目状态", "策划中"
    SetRowValue projectRow, projects, "脚本状态", "未开始"
    SetRowValue projectRow, projects, "优先级", "P1"
    SetRowValue projectRow, projects, "开始日期", Date
    SetRowValue projectRow, projects, "最后更新", Date
    SetRowValue projectRow, projects, "当前脚本页", currentSheetName
    SetRowValue projectRow, projects, "上一版脚本页", previousSheetName
    SetRowValue projectRow, projects, "当前版本", "V1"
    SetRowValue projectRow, projects, "主项目", "否"
    CreateScriptSheets currentSheetName, previousSheetName

    Set scriptIndex = GetTable(SCRIPT_INDEX_TABLE_NAME)
    Set scriptRow = scriptIndex.ListRows.Add
    SetRowValue scriptRow, scriptIndex, "Video ID", videoId
    SetRowValue scriptRow, scriptIndex, "视频标题", videoTitle
    SetRowValue scriptRow, scriptIndex, "当前脚本页", currentSheetName
    SetRowValue scriptRow, scriptIndex, "上一版脚本页", previousSheetName
    SetRowValue scriptRow, scriptIndex, "当前版本", "V1"
    SetRowValue scriptRow, scriptIndex, "脚本状态", "未开始"
    SetRowValue scriptRow, scriptIndex, "最后更新", Date
    CreateVideoForSmokeTest = True
    Exit Function
Failed:
    Debug.Print "Smoke-test video creation failed for " & videoId
    CreateVideoForSmokeTest = False
End Function
