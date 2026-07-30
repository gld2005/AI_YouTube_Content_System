Attribute VB_Name = "modVersioning"
Option Explicit

Public Sub SaveNewVersion()
    Dim videoId As String, reason As String
    videoId = Trim$(InputBox("Enter the Video ID to save:", "Save New Script Version"))
    If Len(videoId) = 0 Then Exit Sub
    reason = Trim$(InputBox("Enter a save reason, or leave blank:", "Save New Script Version"))
    If MsgBox("The current script will overwrite the previous script. Continue?", vbYesNo + vbQuestion, "Save New Script Version") <> vbYes Then Exit Sub
    If Not SaveVersionForVideo(videoId, reason) Then ShowUserError "The current or previous script sheet is missing.", "Version save failed for " & videoId
End Sub

Public Function SaveVersionForVideo(ByVal videoId As String, ByVal reason As String) As Boolean
    Dim projects As ListObject, projectRow As ListRow
    Dim currentSheet As Worksheet, previousSheet As Worksheet
    On Error GoTo Failed
    Set projects = GetTable(PROJECT_TABLE_NAME)
    Set projectRow = FindTableRow(projects, 1, videoId)
    If projectRow Is Nothing Then GoTo Failed
    Set currentSheet = ThisWorkbook.Worksheets(CStr(projectRow.Range.Cells(1, 28).Value))
    Set previousSheet = ThisWorkbook.Worksheets(CStr(projectRow.Range.Cells(1, 29).Value))
    previousSheet.Cells.Clear
    currentSheet.UsedRange.Copy Destination:=previousSheet.Range("A1")
    SetColumnValue projectRow, 31, Date
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn") & " | Version saved for " & videoId & " | " & reason
    SaveVersionForVideo = True
    Exit Function
Failed:
    SaveVersionForVideo = False
End Function
