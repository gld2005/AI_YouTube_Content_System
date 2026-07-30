Attribute VB_Name = "modTasks"
Option Explicit

Public Function RefreshGanttForVideo(ByVal videoId As String) As Boolean
    Dim tasks As ListObject, gantt As ListObject, taskRow As ListRow, ganttRow As ListRow
    On Error GoTo Failed
    Set tasks = GetTable("TasksTable")
    Set gantt = GetTable("GanttTable")
    Do While gantt.ListRows.Count > 0
        gantt.ListRows(1).Delete
    Loop
    For Each taskRow In tasks.ListRows
        If StrComp(CStr(taskRow.Range.Cells(1, 2).Value), videoId, vbTextCompare) = 0 Then
            If IsGanttTask(taskRow) Then
                Set ganttRow = gantt.ListRows.Add
                SetColumnValue ganttRow, 1, taskRow.Range.Cells(1, 1).Value
                SetColumnValue ganttRow, 2, videoId
                SetColumnValue ganttRow, 3, taskRow.Range.Cells(1, 7).Value
                SetColumnValue ganttRow, 4, taskRow.Range.Cells(1, 8).Value
                SetColumnValue ganttRow, 5, taskRow.Range.Cells(1, 20).Value
                SetColumnValue ganttRow, 6, taskRow.Range.Cells(1, 21).Value
                SetColumnValue ganttRow, 7, TaskProgressForGantt(taskRow)
                SetColumnValue ganttRow, 8, DerivedTaskRisk(taskRow)
                SetColumnValue ganttRow, 9, "Yes"
                SetColumnValue ganttRow, 10, taskRow.Range.Cells(1, 26).Value
            End If
        End If
    Next taskRow
    RefreshGanttForVideo = True
    Exit Function
Failed:
    RefreshGanttForVideo = False
End Function

Public Sub RefreshSelectedProjectSchedule()
    Dim videoId As String
    videoId = Trim$(CStr(ActiveCell.Value))
    If Len(videoId) = 0 Then
        ShowUserError "Select a Video ID before refreshing the schedule.", "Schedule refresh was requested with no Video ID."
        Exit Sub
    End If
    If Not RefreshGanttForVideo(videoId) Then
        ShowUserError "The Gantt view could not be refreshed.", "Gantt refresh failed for " & videoId
        Exit Sub
    End If
    MsgBox "The Gantt view was refreshed from main tasks and selected key child tasks.", vbInformation, "Schedule Refresh"
End Sub

Public Function SuggestChildTaskProgress(ByVal parentTaskId As String) As Double
    Dim tasks As ListObject, taskRow As ListRow, totalWeight As Double, completeWeight As Double, weightValue As Double
    On Error GoTo Failed
    Set tasks = GetTable("TasksTable")
    For Each taskRow In tasks.ListRows
        If StrComp(CStr(taskRow.Range.Cells(1, 4).Value), parentTaskId, vbTextCompare) = 0 Then
            weightValue = Val(taskRow.Range.Cells(1, 16).Value)
            If weightValue <= 0 Then weightValue = 1
            totalWeight = totalWeight + weightValue
            If LCase$(Trim$(CStr(taskRow.Range.Cells(1, 10).Value))) = "completed" Then completeWeight = completeWeight + weightValue
        End If
    Next taskRow
    If totalWeight > 0 Then SuggestChildTaskProgress = Round((completeWeight / totalWeight) * 100, 0)
    Exit Function
Failed:
    SuggestChildTaskProgress = 0
End Function

Public Function PlanTasksBackward(ByVal videoId As String) As Boolean
    PlanTasksBackward = PlanTasksBackwardCore(videoId, True)
End Function

Public Function PlanTasksBackwardForSmokeTest(ByVal videoId As String) As Boolean
    PlanTasksBackwardForSmokeTest = PlanTasksBackwardCore(videoId, False)
End Function

Private Function PlanTasksBackwardCore(ByVal videoId As String, ByVal requireConfirmation As Boolean) As Boolean
    Dim projects As ListObject, projectRow As ListRow, tasks As ListObject, taskRow As ListRow
    Dim publishDate As Date, cursorDate As Date, durationDays As Long, rowIndex As Long
    On Error GoTo Failed
    Set projects = GetTable(PROJECT_TABLE_NAME)
    Set projectRow = FindTableRow(projects, 1, videoId)
    If projectRow Is Nothing Then GoTo Failed
    If Not IsDate(projectRow.Range.Cells(1, 17).Value) Then GoTo Failed
    If requireConfirmation Then
        If MsgBox("Reverse planning will update planned start and due dates using natural days. You can adjust the result manually afterward. Continue?", vbYesNo + vbQuestion, "Reverse Planning") <> vbYes Then Exit Function
    End If
    publishDate = CDate(projectRow.Range.Cells(1, 17).Value)
    cursorDate = publishDate
    Set tasks = GetTable("TasksTable")
    For rowIndex = tasks.ListRows.Count To 1 Step -1
        Set taskRow = tasks.ListRows(rowIndex)
        If StrComp(CStr(taskRow.Range.Cells(1, 2).Value), videoId, vbTextCompare) = 0 Then
            durationDays = NaturalTaskDays(taskRow)
            taskRow.Range.Cells(1, 21).Value = cursorDate
            taskRow.Range.Cells(1, 20).Value = DateAdd("d", -(durationDays - 1), cursorDate)
            cursorDate = DateAdd("d", -1, CDate(taskRow.Range.Cells(1, 20).Value))
        End If
    Next rowIndex
    PlanTasksBackwardCore = RefreshGanttForVideo(videoId)
    Exit Function
Failed:
    PlanTasksBackwardCore = False
End Function

Public Sub ShowSelectedTaskDependencyImpact()
    Dim tasks As ListObject, selectedTask As ListRow, taskRow As ListRow, taskId As String, downstreamText As String
    On Error GoTo Failed
    taskId = Trim$(CStr(ActiveCell.Value))
    If Len(taskId) = 0 Then GoTo Failed
    Set tasks = GetTable("TasksTable")
    Set selectedTask = FindTableRow(tasks, 1, taskId)
    If selectedTask Is Nothing Then GoTo Failed
    downstreamText = Trim$(CStr(selectedTask.Range.Cells(1, 25).Value))
    For Each taskRow In tasks.ListRows
        If InStr(1, "," & Replace(CStr(taskRow.Range.Cells(1, 24).Value), " ", "") & ",", "," & taskId & ",", vbTextCompare) > 0 Then
            If Len(downstreamText) > 0 Then downstreamText = downstreamText & ", "
            downstreamText = downstreamText & CStr(taskRow.Range.Cells(1, 1).Value)
        End If
    Next taskRow
    If Len(downstreamText) = 0 Then downstreamText = "None recorded"
    MsgBox "Task: " & taskId & vbCrLf & "Downstream impact: " & downstreamText & vbCrLf & "Publish-date impact: " & CStr(selectedTask.Range.Cells(1, 25).Value), vbInformation, "Dependency Impact"
    Exit Sub
Failed:
    ShowUserError "Select a valid Task ID before viewing dependency impact.", "Dependency impact was requested without a valid task."
End Sub

Private Function IsGanttTask(ByVal taskRow As ListRow) As Boolean
    IsGanttTask = LCase$(Trim$(CStr(taskRow.Range.Cells(1, 5).Value))) = "yes" Or LCase$(Trim$(CStr(taskRow.Range.Cells(1, 6).Value))) = "yes"
End Function

Private Function TaskProgressForGantt(ByVal taskRow As ListRow) As Double
    If LCase$(Trim$(CStr(taskRow.Range.Cells(1, 5).Value))) = "yes" Then
        TaskProgressForGantt = Val(taskRow.Range.Cells(1, 18).Value)
    Else
        TaskProgressForGantt = Val(taskRow.Range.Cells(1, 17).Value)
    End If
End Function

Private Function DerivedTaskRisk(ByVal taskRow As ListRow) As String
    Dim dueDate As Variant, statusValue As String
    statusValue = LCase$(Trim$(CStr(taskRow.Range.Cells(1, 10).Value)))
    If Len(Trim$(CStr(taskRow.Range.Cells(1, 26).Value))) > 0 Then DerivedTaskRisk = "Blocked": Exit Function
    dueDate = taskRow.Range.Cells(1, 21).Value
    If IsDate(dueDate) And statusValue <> "completed" Then
        If CDate(dueDate) < Date Then DerivedTaskRisk = "Overdue": Exit Function
        If DateDiff("d", Date, CDate(dueDate)) <= 3 Then DerivedTaskRisk = "Due Soon": Exit Function
    End If
    DerivedTaskRisk = "Normal"
End Function

Private Function NaturalTaskDays(ByVal taskRow As ListRow) As Long
    If IsDate(taskRow.Range.Cells(1, 20).Value) And IsDate(taskRow.Range.Cells(1, 21).Value) Then
        NaturalTaskDays = DateDiff("d", CDate(taskRow.Range.Cells(1, 20).Value), CDate(taskRow.Range.Cells(1, 21).Value)) + 1
    End If
    If NaturalTaskDays < 1 Then NaturalTaskDays = 1
End Function
