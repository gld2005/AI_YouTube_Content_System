Attribute VB_Name = "modNewVideo"
Option Explicit

Public Sub NewVideoScript()
    Dim videoId As String
    Dim videoTitle As String
    videoId = Trim$(InputBox("Enter a Video ID, for example YT-002:", "New Video Script"))
    If Len(videoId) = 0 Then Exit Sub
    videoTitle = Trim$(InputBox("Enter the video title:", "New Video Script"))
    If Not CreateVideoForSmokeTest(videoId, videoTitle) Then
        ShowUserError "The Video ID already exists or the script pages could not be created.", "New video creation failed for " & videoId
        Exit Sub
    End If
    ThisWorkbook.Worksheets(SafeSheetName(videoId & " - Script")).Activate
    MsgBox "The project and paired script sheets were created.", vbInformation, "New Video Script"
End Sub

Public Function CreateVideoForSmokeTest(ByVal videoId As String, ByVal videoTitle As String) As Boolean
    Dim projects As ListObject, scriptIndex As ListObject
    Dim projectRow As ListRow, scriptRow As ListRow, existingProject As ListRow
    Dim currentSheetName As String, previousSheetName As String
    On Error GoTo Failed
    Set projects = GetTable(PROJECT_TABLE_NAME)
    Set existingProject = FindTableRow(projects, 1, videoId)
    If Not existingProject Is Nothing Then GoTo Failed
    currentSheetName = SafeSheetName(videoId & " - Script")
    previousSheetName = SafeSheetName(videoId & " - Previous")
    If SheetExists(currentSheetName) Or SheetExists(previousSheetName) Then GoTo Failed
    Set projectRow = projects.ListRows.Add
    SetColumnValue projectRow, 1, videoId: SetColumnValue projectRow, 2, videoTitle
    SetColumnValue projectRow, 4, "SER-UNCLASSIFIED": SetColumnValue projectRow, 11, "Planning"
    SetColumnValue projectRow, 13, "Not Started": SetColumnValue projectRow, 14, "P1"
    SetColumnValue projectRow, 16, Date: SetColumnValue projectRow, 28, currentSheetName
    SetColumnValue projectRow, 29, previousSheetName: SetColumnValue projectRow, 30, "V1"
    SetColumnValue projectRow, 31, Date: SetColumnValue projectRow, 32, "No"
    Sheet9.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    ActiveSheet.Name = currentSheetName
    Sheet9.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    ActiveSheet.Name = previousSheetName
    Set scriptIndex = GetTable(SCRIPT_INDEX_TABLE_NAME)
    Set scriptRow = scriptIndex.ListRows.Add
    SetColumnValue scriptRow, 1, videoId: SetColumnValue scriptRow, 2, videoTitle
    SetColumnValue scriptRow, 3, "SER-UNCLASSIFIED": SetColumnValue scriptRow, 4, currentSheetName
    SetColumnValue scriptRow, 5, previousSheetName: SetColumnValue scriptRow, 6, "V1"
    SetColumnValue scriptRow, 7, "Not Started": SetColumnValue scriptRow, 10, Date
    CreateVideoForSmokeTest = True
    Exit Function
Failed:
    Debug.Print "Smoke-test video creation failed for " & videoId
    CreateVideoForSmokeTest = False
End Function
