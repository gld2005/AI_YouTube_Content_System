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

Public Function RunPhase5SmokeTest() As String
    Const testVideoId As String = "TEST-PHASE5-001"
    Const testTitle As String = "Phase 5 Script Controls Smoke Test"
    Dim projects As ListObject
    Dim projectRow As ListRow
    Dim scriptSheet As Worksheet
    Dim types As Variant
    Dim index As Long
    Dim paragraphId As String
    Dim paragraphRow As Long

    On Error GoTo Failed
    If Not CreateVideoForSmokeTest(testVideoId, testTitle) Then GoTo Failed
    Set projects = GetTable(PROJECT_TABLE_NAME)
    Set projectRow = FindTableRow(projects, 1, testVideoId)
    If projectRow Is Nothing Then GoTo Failed
    Set scriptSheet = ThisWorkbook.Worksheets(CStr(projectRow.Range.Cells(1, 28).Value))
    If CStr(scriptSheet.Range("I11").Value) <> "Content" Then GoTo Failed
    If CStr(scriptSheet.Range("I16").Value) <> "Editing" Then GoTo Failed
    If CStr(scriptSheet.Range("I21").Value) <> "Management" Then GoTo Failed
    If CStr(scriptSheet.Range("I26").Value) <> "AI Preview" Then GoTo Failed
    scriptSheet.Activate
    CollapseProductionSidebar
    ExpandProductionSidebar
    AddChapterByValues scriptSheet, "Phase 5 Smoke Test Chapter"
    types = Array("Narration", "On-screen Text", "Quote / Dialogue", "Visual-only", "Music Segment", "Transition")
    For index = LBound(types) To UBound(types)
        paragraphId = AddParagraphByValues(scriptSheet, "Smoke text", "Smoke production note", "Drafting", "00:15")
        paragraphRow = FindParagraphRow(scriptSheet, paragraphId)
        If paragraphRow = 0 Then
            RunPhase5SmokeTest = "FAIL: Paragraph ID was not found: " & paragraphId
            Exit Function
        End If
        ApplyParagraphPresentation scriptSheet, paragraphRow, CStr(types(index)), "Drafting"
        If CStr(scriptSheet.Range("B" & paragraphRow).Value) <> CStr(types(index)) Then GoTo Failed
    Next index
    If NormalizeParagraphStatus("Completed") <> "Completed" Then GoTo Failed
    If NormalizeParagraphType("Transition") <> "Transition" Then GoTo Failed
    RunPhase5SmokeTest = "PASS: story flow, sidebar groups, paragraph types, and statuses are available."
    Exit Function
Failed:
    RunPhase5SmokeTest = "FAIL: Phase 5 script controls smoke test did not complete. " & Err.Description
End Function

Public Function RunPhase6SmokeTest() As String
    Const testVideoId As String = "TEST-PHASE6-001"
    Const testTitle As String = "Phase 6 Version and Recycle Smoke Test"
    Dim projects As ListObject, projectRow As ListRow, scriptSheet As Worksheet
    Dim firstId As String, secondId As String, firstRow As Long, secondRow As Long, recycleId As String
    Dim stage As String

    On Error GoTo Failed
    stage = "create"
    If Not CreateVideoForSmokeTest(testVideoId, testTitle) Then GoTo Failed
    stage = "project"
    Set projects = GetTable(PROJECT_TABLE_NAME)
    Set projectRow = FindTableRow(projects, 1, testVideoId)
    If projectRow Is Nothing Then GoTo Failed
    Set scriptSheet = ThisWorkbook.Worksheets(CStr(projectRow.Range.Cells(1, 28).Value))
    stage = "chapter"
    AddChapterByValues scriptSheet, "Phase 6 Smoke Test Chapter"
    stage = "first paragraph"
    firstId = AddParagraphByValues(scriptSheet, "Original narration.", "Original visual note.", "Drafting", "00:15")
    stage = "second paragraph"
    secondId = AddParagraphByValues(scriptSheet, "Unchanged narration.", "Unchanged visual note.", "Drafting", "00:15")
    stage = "baseline save"
    If Not SaveVersionForVideo(testVideoId, "Baseline") Then
        RunPhase6SmokeTest = "FAIL: Baseline version save failed. " & LastVersionError
        Exit Function
    End If
    stage = "difference"
    firstRow = FindParagraphRow(scriptSheet, firstId)
    scriptSheet.Cells(firstRow + 1, "B").Value = "Changed narration."
    If Not SaveVersionForVideo(testVideoId, "Changed narration") Then GoTo Failed
    If scriptSheet.Cells(firstRow + 1, "B").Interior.Color <> RGB(237, 204, 188) Then GoTo Failed
    stage = "paragraph delete"
    secondRow = FindParagraphRow(scriptSheet, secondId)
    scriptSheet.Activate
    scriptSheet.Cells(secondRow, "A").Select
    DeleteSelectedParagraph
    recycleId = NewestRecycleItemId(scriptSheet)
    If Len(recycleId) = 0 Then GoTo Failed
    If Not RestoreRecycleItem(scriptSheet, recycleId) Then GoTo Failed
    stage = "paragraph restore"
    If FindParagraphRow(scriptSheet, secondId) = 0 Then GoTo Failed
    scriptSheet.Activate
    scriptSheet.Cells(14, "A").Select
    stage = "chapter delete"
    DeleteSelectedChapter
    recycleId = NewestRecycleItemId(scriptSheet)
    If Len(recycleId) = 0 Then GoTo Failed
    stage = "chapter restore"
    If Not RestoreRecycleItem(scriptSheet, recycleId) Then GoTo Failed
    If Left$(CStr(scriptSheet.Cells(14, "A").Value), 3) <> "CH-" Then GoTo Failed
    RunPhase6SmokeTest = "PASS: version differences and restorable recycle items are working."
    Exit Function
Failed:
    RunPhase6SmokeTest = "FAIL: Phase 6 version and recycle smoke test did not complete at " & stage & ". " & Err.Description
End Function
