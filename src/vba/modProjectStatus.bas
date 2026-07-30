Attribute VB_Name = "modProjectStatus"
Option Explicit

Private Const PROJECT_TOTAL_PROGRESS_COLUMN As Long = 19
Private Const PROJECT_STRUCTURE_PROGRESS_COLUMN As Long = 20
Private Const PROJECT_WRITING_PROGRESS_COLUMN As Long = 21
Private Const PROJECT_RECORDING_PROGRESS_COLUMN As Long = 22
Private Const PROJECT_ASSET_PROGRESS_COLUMN As Long = 23
Private Const PROJECT_EDITING_PROGRESS_COLUMN As Long = 24
Private Const PROJECT_REMAINING_DAYS_COLUMN As Long = 25
Private Const PROJECT_RISK_COLUMN As Long = 26

Public Function RefreshProjectProgress(ByVal videoId As String) As Boolean
    Dim projects As ListObject, projectRow As ListRow, scriptSheet As Worksheet
    Dim structureProgress As Double, writingProgress As Double
    Dim recordingProgress As Double, assetProgress As Double, editingProgress As Double
    On Error GoTo Failed
    Set projects = GetTable(PROJECT_TABLE_NAME)
    Set projectRow = FindTableRow(projects, 1, videoId)
    If projectRow Is Nothing Then GoTo Failed
    Set scriptSheet = ThisWorkbook.Worksheets(CStr(projectRow.Range.Cells(1, 28).Value))
    structureProgress = CalculateStructureProgress(scriptSheet)
    writingProgress = CalculateWritingProgress(scriptSheet)
    recordingProgress = CalculateTaskStageProgress(videoId, "Recording")
    assetProgress = CalculateTaskStageProgress(videoId, "Assets")
    editingProgress = CalculateTaskStageProgress(videoId, "Edit")
    SetColumnValue projectRow, PROJECT_STRUCTURE_PROGRESS_COLUMN, structureProgress
    SetColumnValue projectRow, PROJECT_WRITING_PROGRESS_COLUMN, writingProgress
    SetColumnValue projectRow, PROJECT_RECORDING_PROGRESS_COLUMN, recordingProgress
    SetColumnValue projectRow, PROJECT_ASSET_PROGRESS_COLUMN, assetProgress
    SetColumnValue projectRow, PROJECT_EDITING_PROGRESS_COLUMN, editingProgress
    SetColumnValue projectRow, PROJECT_TOTAL_PROGRESS_COLUMN, Round((structureProgress * 0.2) + (writingProgress * 0.35) + (recordingProgress * 0.15) + (assetProgress * 0.15) + (editingProgress * 0.15), 0)
    SetColumnValue projectRow, PROJECT_REMAINING_DAYS_COLUMN, RemainingDays(projectRow.Range.Cells(1, 17).Value)
    SetColumnValue projectRow, PROJECT_RISK_COLUMN, ProjectRisk(projectRow.Range.Cells(1, 17).Value, projectRow.Range.Cells(1, PROJECT_TOTAL_PROGRESS_COLUMN).Value)
    RefreshProjectProgress = True
    Exit Function
Failed:
    RefreshProjectProgress = False
End Function

Public Sub SuggestSelectedProjectStage()
    Dim projectRow As ListRow, projects As ListObject, videoId As String, suggestedStage As String
    Set projects = GetTable(PROJECT_TABLE_NAME)
    Set projectRow = FindTableRow(projects, 1, CStr(ActiveCell.Value))
    If projectRow Is Nothing Then
        ShowUserError "Select a Video ID in the project table first.", "Project-stage suggestion was requested without a project row."
        Exit Sub
    End If
    videoId = CStr(projectRow.Range.Cells(1, 1).Value)
    If Not RefreshProjectProgress(videoId) Then
        ShowUserError "Project progress could not be refreshed.", "Project progress refresh failed for " & videoId
        Exit Sub
    End If
    suggestedStage = SuggestedStage(CDbl(projectRow.Range.Cells(1, PROJECT_TOTAL_PROGRESS_COLUMN).Value), CStr(projectRow.Range.Cells(1, PROJECT_RISK_COLUMN).Value))
    MsgBox "Suggested working stage: " & suggestedStage & vbCrLf & "No project status or stage has been changed. Confirm any transition manually.", vbInformation, "Project Stage Suggestion"
End Sub

Private Function CalculateStructureProgress(ByVal ws As Worksheet) As Double
    Dim rowIndex As Long, chapterCount As Long, completeCount As Long, hasFields As Boolean
    For rowIndex = 12 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If Left$(CStr(ws.Cells(rowIndex, "A").Value), 3) = "CH-" Then
            chapterCount = chapterCount + 1
            hasFields = Len(Trim$(CStr(ws.Cells(rowIndex + 1, "B").Value))) > 0 And Len(Trim$(CStr(ws.Cells(rowIndex + 2, "B").Value))) > 0 And Len(Trim$(CStr(ws.Cells(rowIndex + 3, "B").Value))) > 0 And Len(Trim$(CStr(ws.Cells(rowIndex + 4, "B").Value))) > 0
            If hasFields And LCase$(Trim$(CStr(ws.Cells(rowIndex, "G").Value))) = "confirmed" Then completeCount = completeCount + 1
        End If
    Next rowIndex
    If chapterCount > 0 Then CalculateStructureProgress = Round((completeCount / chapterCount) * 100, 0)
End Function

Private Function CalculateWritingProgress(ByVal ws As Worksheet) As Double
    Dim rowIndex As Long, totalWeight As Double, completeWeight As Double, durationWeight As Double, statusValue As String
    For rowIndex = 12 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If Len(CStr(ws.Cells(rowIndex, "A").Value)) = 5 And Mid$(CStr(ws.Cells(rowIndex, "A").Value), 3, 1) = "-" Then
            durationWeight = DurationWeight(CStr(ws.Cells(rowIndex, "D").Value))
            totalWeight = totalWeight + durationWeight
            statusValue = CStr(ws.Cells(rowIndex, "C").Value)
            completeWeight = completeWeight + (durationWeight * WritingStatusWeight(statusValue))
        End If
    Next rowIndex
    If totalWeight > 0 Then CalculateWritingProgress = Round((completeWeight / totalWeight) * 100, 0)
End Function

Private Function CalculateTaskStageProgress(ByVal videoId As String, ByVal stageName As String) As Double
    Dim tasks As ListObject, taskRow As ListRow, totalWeight As Double, completeWeight As Double, weightValue As Double
    On Error GoTo NoTasks
    Set tasks = GetTable("TasksTable")
    For Each taskRow In tasks.ListRows
        If StrComp(CStr(taskRow.Range.Cells(1, 2).Value), videoId, vbTextCompare) = 0 And StrComp(CStr(taskRow.Range.Cells(1, 7).Value), stageName, vbTextCompare) = 0 Then
            weightValue = Val(taskRow.Range.Cells(1, 16).Value)
            If weightValue <= 0 Then weightValue = 1
            totalWeight = totalWeight + weightValue
            completeWeight = completeWeight + (weightValue * TaskStatusWeight(CStr(taskRow.Range.Cells(1, 10).Value)))
        End If
    Next taskRow
    If totalWeight > 0 Then CalculateTaskStageProgress = Round((completeWeight / totalWeight) * 100, 0)
NoTasks:
End Function

Private Function DurationWeight(ByVal durationText As String) As Double
    Dim values As Variant
    values = Split(durationText, ":")
    If UBound(values) = 1 And IsNumeric(values(0)) And IsNumeric(values(1)) Then DurationWeight = (CDbl(values(0)) * 60) + CDbl(values(1))
    If DurationWeight <= 0 Then DurationWeight = 1
End Function

Private Function WritingStatusWeight(ByVal statusValue As String) As Double
    Select Case LCase$(Trim$(statusValue))
        Case "completed": WritingStatusWeight = 1
        Case "edited", "assets matched", "recorded": WritingStatusWeight = 0.9
        Case "finalized": WritingStatusWeight = 0.75
        Case "drafting": WritingStatusWeight = 0.45
        Case Else: WritingStatusWeight = 0
    End Select
End Function

Private Function TaskStatusWeight(ByVal statusValue As String) As Double
    If LCase$(Trim$(statusValue)) = "completed" Then TaskStatusWeight = 1
End Function

Private Function RemainingDays(ByVal targetDate As Variant) As Variant
    If IsDate(targetDate) Then RemainingDays = DateDiff("d", Date, CDate(targetDate)) Else RemainingDays = ""
End Function

Private Function ProjectRisk(ByVal targetDate As Variant, ByVal progressValue As Variant) As String
    If Not IsDate(targetDate) Then Exit Function
    If DateDiff("d", Date, CDate(targetDate)) < 0 And Val(progressValue) < 100 Then
        ProjectRisk = "Overdue"
    ElseIf DateDiff("d", Date, CDate(targetDate)) <= 3 And Val(progressValue) < 85 Then
        ProjectRisk = "Due Soon"
    Else
        ProjectRisk = "Normal"
    End If
End Function

Private Function SuggestedStage(ByVal totalProgress As Double, ByVal riskValue As String) As String
    If LCase$(riskValue) = "blocked" Then SuggestedStage = "No change: blocked": Exit Function
    If totalProgress >= 100 Then SuggestedStage = "Optimize" ElseIf totalProgress >= 90 Then SuggestedStage = "Publish" ElseIf totalProgress >= 85 Then SuggestedStage = "Post" ElseIf totalProgress >= 60 Then SuggestedStage = "Edit" Else SuggestedStage = "Script"
End Function
