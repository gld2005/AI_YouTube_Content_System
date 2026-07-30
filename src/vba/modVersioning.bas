Attribute VB_Name = "modVersioning"
Option Explicit

Public LastVersionError As String

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
    Dim currentSheet As Worksheet, previousSheet As Worksheet, replacementSheet As Worksheet
    Dim operation As String
    Dim previousSheetName As String, temporarySheetName As String
    On Error GoTo Failed
    LastVersionError = ""
    operation = "locate project table"
    Set projects = GetTable(PROJECT_TABLE_NAME)
    Set projectRow = FindTableRow(projects, 1, videoId)
    If projectRow Is Nothing Then GoTo Failed
    operation = "locate current script"
    Set currentSheet = ThisWorkbook.Worksheets(CStr(projectRow.Range.Cells(1, 28).Value))
    operation = "locate previous script"
    Set previousSheet = ThisWorkbook.Worksheets(CStr(projectRow.Range.Cells(1, 29).Value))
    previousSheetName = previousSheet.Name
    operation = "compare paragraphs"
    MarkChangedParagraphs currentSheet, previousSheet
    operation = "copy current script to replacement"
    currentSheet.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set replacementSheet = ActiveSheet
    temporarySheetName = SafeSheetName("TMP-" & Format$(Now, "yymmddhhnnss"))
    replacementSheet.Name = temporarySheetName
    operation = "rename existing previous script"
    previousSheet.Name = SafeSheetName("OLD-" & Format$(Now, "yymmddhhnnss"))
    operation = "publish replacement script"
    replacementSheet.Name = previousSheetName
    operation = "remove replaced previous script"
    Application.DisplayAlerts = False
    previousSheet.Delete
    Application.DisplayAlerts = True
    SetColumnValue projectRow, 31, Date
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn") & " | Version saved for " & videoId & " | " & reason
    SaveVersionForVideo = True
    Exit Function
Failed:
    Application.DisplayAlerts = True
    LastVersionError = operation & ": " & Err.Description
    Debug.Print "Version save failed: " & Err.Number & " | " & LastVersionError
    SaveVersionForVideo = False
End Function

Private Sub MarkChangedParagraphs(ByVal currentSheet As Worksheet, ByVal previousSheet As Worksheet)
    Dim cell As Range, currentRow As Long, previousRow As Long, lastRow As Long
    lastRow = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
    For Each cell In currentSheet.Range("A12:A" & lastRow).Cells
        If IsParagraphId(CStr(cell.Value)) Then
            currentRow = cell.Row
            previousRow = FindParagraphRow(previousSheet, CStr(cell.Value))
            currentSheet.Range("B" & currentRow + 1).Interior.Color = RGB(250, 247, 241)
            currentSheet.Range("B" & currentRow + 2).Interior.Color = RGB(250, 247, 241)
            If previousRow = 0 Then
                currentSheet.Range("B" & currentRow + 1).Interior.Color = RGB(237, 204, 188)
                currentSheet.Range("B" & currentRow + 2).Interior.Color = RGB(237, 204, 188)
            ElseIf ParagraphSignature(currentSheet, currentRow) <> ParagraphSignature(previousSheet, previousRow) Then
                currentSheet.Range("B" & currentRow + 1).Interior.Color = RGB(237, 204, 188)
                currentSheet.Range("B" & currentRow + 2).Interior.Color = RGB(237, 204, 188)
            End If
        End If
    Next cell
End Sub

Private Function ParagraphSignature(ByVal ws As Worksheet, ByVal paragraphRow As Long) As String
    ParagraphSignature = CStr(ws.Cells(paragraphRow, "B").Value) & "|" & CStr(ws.Cells(paragraphRow, "C").Value) & "|" & CStr(ws.Cells(paragraphRow, "D").Value) & "|" & CStr(ws.Cells(paragraphRow + 1, "B").Value) & "|" & CStr(ws.Cells(paragraphRow + 2, "B").Value)
End Function

Private Function IsParagraphId(ByVal value As String) As Boolean
    IsParagraphId = Len(value) = 5 And Mid$(value, 3, 1) = "-"
End Function
