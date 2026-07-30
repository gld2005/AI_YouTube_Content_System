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
