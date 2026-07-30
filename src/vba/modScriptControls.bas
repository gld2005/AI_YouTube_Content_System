Attribute VB_Name = "modScriptControls"
Option Explicit

Public Function PromptParagraphType() As String
    Dim value As String
    value = Trim$(InputBox("Choose a paragraph type: Narration, On-screen Text, Quote / Dialogue, Visual-only, Music Segment, or Transition.", "Paragraph Type", "Narration"))
    PromptParagraphType = NormalizeParagraphType(value)
End Function

Public Function PromptParagraphStatus() As String
    Dim value As String
    value = Trim$(InputBox("Choose a paragraph status: To Ideate, Drafting, Finalized, Recorded, Assets Matched, Edited, or Completed.", "Paragraph Status", "Drafting"))
    PromptParagraphStatus = NormalizeParagraphStatus(value)
End Function

Public Function NormalizeParagraphType(ByVal value As String) As String
    Select Case LCase$(Trim$(value))
        Case "narration": NormalizeParagraphType = "Narration"
        Case "on-screen text": NormalizeParagraphType = "On-screen Text"
        Case "quote / dialogue": NormalizeParagraphType = "Quote / Dialogue"
        Case "visual-only": NormalizeParagraphType = "Visual-only"
        Case "music segment": NormalizeParagraphType = "Music Segment"
        Case "transition": NormalizeParagraphType = "Transition"
        Case Else: NormalizeParagraphType = "Narration"
    End Select
End Function

Public Function NormalizeParagraphStatus(ByVal value As String) As String
    Select Case LCase$(Trim$(value))
        Case "to ideate": NormalizeParagraphStatus = "To Ideate"
        Case "drafting", "draft": NormalizeParagraphStatus = "Drafting"
        Case "finalized": NormalizeParagraphStatus = "Finalized"
        Case "recorded": NormalizeParagraphStatus = "Recorded"
        Case "assets matched": NormalizeParagraphStatus = "Assets Matched"
        Case "edited": NormalizeParagraphStatus = "Edited"
        Case "completed": NormalizeParagraphStatus = "Completed"
        Case Else: NormalizeParagraphStatus = "Drafting"
    End Select
End Function

Public Sub ApplyParagraphPresentation(ByVal ws As Worksheet, ByVal paragraphRow As Long, ByVal paragraphType As String, ByVal paragraphStatus As String)
    Dim typeColor As Long
    Dim statusColor As Long

    paragraphType = NormalizeParagraphType(paragraphType)
    paragraphStatus = NormalizeParagraphStatus(paragraphStatus)
    ws.Range("B" & paragraphRow).Value = paragraphType
    ws.Range("C" & paragraphRow).Value = paragraphStatus
    typeColor = ParagraphTypeColor(paragraphType)
    statusColor = ParagraphStatusColor(paragraphStatus)
    ws.Range("A" & paragraphRow & ":A" & paragraphRow + 2).Interior.Color = typeColor
    ws.Range("A" & paragraphRow).Font.Color = RGB(255, 255, 255)
    ws.Range("B" & paragraphRow).Interior.Color = typeColor
    ws.Range("C" & paragraphRow).Interior.Color = statusColor
    AddParagraphValidation ws, paragraphRow
End Sub

Public Sub RefreshSelectedParagraphPresentation()
    Dim paragraphRow As Long
    paragraphRow = SelectedParagraphRow(ActiveSheet, ActiveCell.Row)
    If paragraphRow = 0 Then
        ShowUserError "Select a paragraph block before refreshing its style.", "No paragraph ID was found above the selected row."
        Exit Sub
    End If
    ApplyParagraphPresentation ActiveSheet, paragraphRow, CStr(ActiveSheet.Range("B" & paragraphRow).Value), CStr(ActiveSheet.Range("C" & paragraphRow).Value)
End Sub

Public Sub CollapseProductionSidebar()
    ActiveSheet.Outline.ShowLevels ColumnLevels:=1
End Sub

Public Sub ExpandProductionSidebar()
    ActiveSheet.Outline.ShowLevels ColumnLevels:=2
End Sub

Public Function SelectedParagraphRow(ByVal ws As Worksheet, ByVal selectedRow As Long) As Long
    Dim scanRow As Long
    Dim paragraphId As String
    For scanRow = selectedRow To 12 Step -1
        paragraphId = CStr(ws.Cells(scanRow, "A").Value)
        If Len(paragraphId) = 5 And Mid$(paragraphId, 3, 1) = "-" Then
            SelectedParagraphRow = scanRow
            Exit Function
        End If
    Next scanRow
End Function

Public Function FindParagraphRow(ByVal ws As Worksheet, ByVal paragraphId As String) As Long
    Dim matchCell As Range
    Set matchCell = ws.Columns("A").Find(What:=paragraphId, After:=ws.Range("A1"), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If Not matchCell Is Nothing Then FindParagraphRow = matchCell.Row
End Function

Private Sub AddParagraphValidation(ByVal ws As Worksheet, ByVal paragraphRow As Long)
    On Error Resume Next
    ws.Range("B" & paragraphRow).Validation.Delete
    ws.Range("B" & paragraphRow).Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "Narration,On-screen Text,Quote / Dialogue,Visual-only,Music Segment,Transition"
    ws.Range("C" & paragraphRow).Validation.Delete
    ws.Range("C" & paragraphRow).Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "To Ideate,Drafting,Finalized,Recorded,Assets Matched,Edited,Completed"
    On Error GoTo 0
End Sub

Private Function ParagraphTypeColor(ByVal paragraphType As String) As Long
    Select Case paragraphType
        Case "Narration": ParagraphTypeColor = RGB(183, 207, 220)
        Case "On-screen Text": ParagraphTypeColor = RGB(211, 196, 220)
        Case "Quote / Dialogue": ParagraphTypeColor = RGB(220, 214, 198)
        Case "Visual-only": ParagraphTypeColor = RGB(184, 214, 206)
        Case "Music Segment": ParagraphTypeColor = RGB(224, 205, 190)
        Case "Transition": ParagraphTypeColor = RGB(201, 211, 218)
    End Select
End Function

Private Function ParagraphStatusColor(ByVal paragraphStatus As String) As Long
    Select Case paragraphStatus
        Case "To Ideate": ParagraphStatusColor = RGB(224, 220, 211)
        Case "Drafting", "Finalized": ParagraphStatusColor = RGB(198, 215, 227)
        Case "Recorded", "Assets Matched", "Edited": ParagraphStatusColor = RGB(230, 205, 191)
        Case "Completed": ParagraphStatusColor = RGB(196, 216, 194)
    End Select
End Function
