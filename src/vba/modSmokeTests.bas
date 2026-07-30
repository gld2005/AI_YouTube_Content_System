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
