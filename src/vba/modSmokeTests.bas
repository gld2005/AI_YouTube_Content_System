Attribute VB_Name = "modSmokeTests"
Option Explicit

Public Function RunPhase3SmokeTest() As String
    Const testVideoId As String = "TEST-SMOKE-001"
    Const testTitle As String = "Phase 3 Macro Smoke Test"
    Dim projects As ListObject
    Dim projectRow As ListRow
    Dim currentSheetName As String
    Dim previousSheetName As String

    On Error GoTo Failed
    If Not CreateVideoForSmokeTest(testVideoId, testTitle) Then GoTo Failed
    Set projects = GetTable(PROJECT_TABLE_NAME)
    Set projectRow = FindTableRow(projects, 1, testVideoId)
    If projectRow Is Nothing Then GoTo Failed
    currentSheetName = CStr(projectRow.Range.Cells(1, 28).Value)
    previousSheetName = CStr(projectRow.Range.Cells(1, 29).Value)
    If Not SheetExists(currentSheetName) Or Not SheetExists(previousSheetName) Then GoTo Failed
    RunPhase3SmokeTest = "PASS: project record, script index, and paired script pages created."
    Exit Function
Failed:
    RunPhase3SmokeTest = "FAIL: Phase 3 macro smoke test did not complete."
End Function

Public Function RunPhase4SmokeTest() As String
    Const testVideoId As String = "TEST-PHASE4-001"
    Const testTitle As String = "Phase 4 Script Workspace Smoke Test"
    Dim projects As ListObject
    Dim projectRow As ListRow
    Dim scriptSheet As Worksheet
    Dim chapterRow As Long
    Dim paragraphId As String

    On Error GoTo Failed
    If Not CreateVideoForSmokeTest(testVideoId, testTitle) Then GoTo Failed
    Set projects = GetTable(PROJECT_TABLE_NAME)
    Set projectRow = FindTableRow(projects, 1, testVideoId)
    If projectRow Is Nothing Then GoTo Failed
    Set scriptSheet = ThisWorkbook.Worksheets(CStr(projectRow.Range.Cells(1, 28).Value))
    If InStr(1, CStr(scriptSheet.Range("A1").Value), "Script Workspace", vbTextCompare) = 0 Then GoTo Failed
    chapterRow = AddChapterByValues(scriptSheet, "Smoke Test Chapter")
    If chapterRow < 12 Then GoTo Failed
    paragraphId = AddParagraphByValues(scriptSheet, "Smoke test narration.", "Smoke test visual note.", "Draft", "00:30")
    If paragraphId <> "01-01" Then GoTo Failed
    RunPhase4SmokeTest = "PASS: English script workspace, chapter, and three-row paragraph block created."
    Exit Function
Failed:
    RunPhase4SmokeTest = "FAIL: Phase 4 script workspace smoke test did not complete."
End Function
