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

Public Function RunPhase7SmokeTest() As String
    Const testVideoId As String = "TEST-PHASE7-001"
    Const testTitle As String = "Phase 7 Relationship Smoke Test"
    Dim sources As ListObject, evidence As ListObject, sourceRow As ListRow, evidenceRow As ListRow
    Dim projects As ListObject, projectRow As ListRow, scriptSheet As Worksheet
    Dim paragraphId As String, paragraphRow As Long, chapterId As String

    On Error GoTo Failed
    If Not CreateVideoForSmokeTest(testVideoId, testTitle) Then GoTo Failed
    Set sources = GetTable(SOURCE_TABLE_NAME)
    Set sourceRow = sources.ListRows.Add
    SetColumnValue sourceRow, 1, "SRC-PHASE7-001": SetColumnValue sourceRow, 3, "Smoke Test Source"
    Set evidence = GetTable(EVIDENCE_TABLE_NAME)
    Set evidenceRow = evidence.ListRows.Add
    SetColumnValue evidenceRow, 1, "EVD-PHASE7-001": SetColumnValue evidenceRow, 2, "SRC-PHASE7-001"
    SetColumnValue evidenceRow, 3, "Verifiable Fact": SetColumnValue evidenceRow, 4, "Smoke test evidence."
    Set projects = GetTable(PROJECT_TABLE_NAME)
    Set projectRow = FindTableRow(projects, 1, testVideoId)
    Set scriptSheet = ThisWorkbook.Worksheets(CStr(projectRow.Range.Cells(1, 28).Value))
    AddChapterByValues scriptSheet, "Phase 7 Smoke Test Chapter"
    paragraphId = AddParagraphByValues(scriptSheet, "Linked narration.", "Linked production note.", "Drafting", "00:15")
    paragraphRow = FindParagraphRow(scriptSheet, paragraphId)
    chapterId = "CH-01"
    If Not AddResearchLink(testVideoId, "EVD-PHASE7-001", chapterId, paragraphId, "Script evidence") Then GoTo Failed
    If Not AddPackagingConcept(testVideoId, "Smoke Test Title", "Smoke Thumbnail") Then GoTo Failed
    If Not AddProjectAssetLink(testVideoId, chapterId, paragraphId, "Smoke Visual Asset") Then GoTo Failed
    RunPhase7SmokeTest = "PASS: source, evidence, packaging, asset, and paragraph links were created."
    Exit Function
Failed:
    RunPhase7SmokeTest = "FAIL: Phase 7 relationship smoke test did not complete. " & Err.Description
End Function

Public Function RunPhase8SmokeTest() As String
    Const testVideoId As String = "TEST-PHASE8-001"
    Const testTitle As String = "Phase 8 Approval Smoke Test"
    Dim projects As ListObject, projectRow As ListRow, scriptSheet As Worksheet
    Dim paragraphId As String, paragraphRow As Long, suggestionId As String
    Dim stage As String

    On Error GoTo Failed
    stage = "create"
    If Not CreateVideoForSmokeTest(testVideoId, testTitle) Then GoTo Failed
    Set projects = GetTable(PROJECT_TABLE_NAME)
    Set projectRow = FindTableRow(projects, 1, testVideoId)
    Set scriptSheet = ThisWorkbook.Worksheets(CStr(projectRow.Range.Cells(1, 28).Value))
    stage = "paragraph"
    AddChapterByValues scriptSheet, "Phase 8 Smoke Test Chapter"
    paragraphId = AddParagraphByValues(scriptSheet, "Original approved text.", "", "Drafting", "00:15")
    paragraphRow = FindParagraphRow(scriptSheet, paragraphId)
    stage = "queue"
    suggestionId = QueueAiSuggestion(testVideoId, "CH-01", paragraphId, "Polish Current Paragraph", "Original approved text.", "Candidate text awaiting approval.")
    If Len(suggestionId) = 0 Then GoTo Failed
    stage = "pending check"
    If CStr(scriptSheet.Cells(paragraphRow + 1, "B").Value) <> "Original approved text." Then GoTo Failed
    stage = "approval"
    If Not ApproveFullAdoption(suggestionId, scriptSheet) Then
        RunPhase8SmokeTest = "FAIL: Approval failed. " & LastApprovalError
        Exit Function
    End If
    If CStr(scriptSheet.Cells(paragraphRow + 1, "B").Value) <> "Candidate text awaiting approval." Then GoTo Failed
    stage = "publish review"
    If Not AddPublishReview(testVideoId, "YouTube", "Smoke Test Final Title") Then GoTo Failed
    RunPhase8SmokeTest = "PASS: AI suggestion approval is gated and publish review was recorded."
    Exit Function
Failed:
    RunPhase8SmokeTest = "FAIL: Phase 8 approval smoke test did not complete at " & stage & ". " & Err.Description
End Function
