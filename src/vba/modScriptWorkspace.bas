Attribute VB_Name = "modScriptWorkspace"
Option Explicit

Private Const CONTENT_START_ROW As Long = 12

Public Sub InitializeScriptWorkspace(ByVal ws As Worksheet, ByVal videoId As String, ByVal videoTitle As String)
    Application.ScreenUpdating = False
    ws.Cells.Clear
    ws.Cells.UnMerge
    ws.Cells.Font.Name = "Aptos"
    ws.Cells.Font.Size = 10
    ws.Cells.Interior.Color = RGB(250, 247, 241)
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B:G").ColumnWidth = 13
    ws.Columns("H").ColumnWidth = 3
    ws.Columns("I:K").ColumnWidth = 15
    ws.Columns("I:K").Columns.Group
    ws.Outline.SummaryColumn = xlRight
    With ws.Range("A1:K1")
        .Merge
        .Value = "Script Workspace | " & videoId & " | " & videoTitle
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(73, 100, 119)
    End With
    ws.Rows(1).RowHeight = 28
    WriteSummary ws, "A3", "Core Promise", "Define the audience value before drafting."
    WriteSummary ws, "A5", "Selected Title", videoTitle
    WriteSummary ws, "D3", "Hook", "Open with the strongest tension or question."
    WriteSummary ws, "D5", "Target Duration", "00:00"
    WriteSummary ws, "I3", "Current Stage", "Planning"
    WriteSummary ws, "I5", "Version", "V1"
    WriteBand ws, "A9:G9", "STORY FLOW AND CHAPTER OUTLINE", RGB(102, 132, 154)
    WriteStoryFlow ws
    WriteBand ws, "I9:K9", "PRODUCTION SIDEBAR", RGB(172, 117, 94)
    WriteSidebarGroup ws, 11, "Content", "Visual idea, Source IDs, and asset requirements."
    WriteSidebarGroup ws, 16, "Editing", "Subtitle emphasis, music, animation, and timecodes."
    WriteSidebarGroup ws, 21, "Management", "Status, change reason, and last update."
    WriteSidebarGroup ws, 26, "AI Preview", "Suggestions stay pending until explicit approval."
    WriteBand ws, "I32:K32", "Script Actions", RGB(230, 221, 207)
    ws.Range("I33").Value = "Add Chapter": ws.Range("J33").Value = "Add Paragraph": ws.Range("K33").Value = "Save Version"
    ws.Range("I33:K33").Interior.Color = RGB(218, 229, 236)
    ws.Range("I33:K33").Font.Bold = True
    ws.Range("A" & CONTENT_START_ROW).Value = "Use Add Chapter to begin the script outline."
    ws.Range("A" & CONTENT_START_ROW).Font.Italic = True
    ws.Range("A" & CONTENT_START_ROW).Font.Color = RGB(120, 120, 120)
    ws.Range("A:K").VerticalAlignment = xlTop
    ws.Outline.ShowLevels ColumnLevels:=2
    Application.ScreenUpdating = True
End Sub

Public Sub AddChapter()
    Dim chapterTitle As String
    chapterTitle = Trim$(InputBox("Enter a chapter title:", "Add Chapter"))
    If Len(chapterTitle) > 0 Then AddChapterByValues ActiveSheet, chapterTitle
End Sub

Public Function AddChapterByValues(ByVal ws As Worksheet, ByVal chapterTitle As String) As Long
    Dim chapterRow As Long, chapterNumber As Long
    chapterNumber = CountPrefix(ws, "CH-") + 1
    chapterRow = NextContentRow(ws)
    If chapterRow < CONTENT_START_ROW Then chapterRow = CONTENT_START_ROW
    ws.Rows(chapterRow & ":" & chapterRow + 4).Insert Shift:=xlDown
    With ws.Range("A" & chapterRow & ":G" & chapterRow)
        .Merge
        .Value = "CH-" & Format$(chapterNumber, "00") & " | " & chapterTitle
        .Font.Bold = True
        .Font.Color = RGB(35, 58, 75)
        .Interior.Color = RGB(218, 229, 236)
    End With
    WriteChapterField ws, chapterRow + 1, "Objective", "Define the chapter outcome."
    WriteChapterField ws, chapterRow + 2, "Main Claim", "State the single claim this chapter must establish."
    WriteChapterField ws, chapterRow + 3, "Subclaims", "Add two to four supporting claims."
    WriteChapterField ws, chapterRow + 4, "Transition", "Write the bridge to the next chapter."
    AddChapterByValues = chapterRow
End Function

Public Sub AddParagraph()
    Dim narration As String
    Dim paragraphRow As Long
    narration = Trim$(InputBox("Enter the narration or leave it blank:", "Add Paragraph"))
    paragraphRow = SelectedParagraphInsertRow(ActiveSheet, ActiveCell.Row)
    If paragraphRow = 0 Then
        ShowUserError "Select an existing paragraph before adding a new paragraph.", "Add Paragraph requires a selected paragraph block."
        Exit Sub
    End If
    AddParagraphAtRow ActiveSheet, paragraphRow, narration, "", PromptParagraphStatus(), "00:00", PromptParagraphType()
End Sub

Public Function AddParagraphByValues(ByVal ws As Worksheet, ByVal narration As String, ByVal productionNotes As String, ByVal paragraphStatus As String, ByVal estimatedDuration As String) As String
    Dim paragraphRow As Long
    paragraphRow = NextContentRow(ws)
    If paragraphRow < CONTENT_START_ROW Then paragraphRow = CONTENT_START_ROW
    AddParagraphByValues = AddParagraphAtRow(ws, paragraphRow, narration, productionNotes, paragraphStatus, estimatedDuration, "Narration")
End Function

Private Function AddParagraphAtRow(ByVal ws As Worksheet, ByVal paragraphRow As Long, ByVal narration As String, ByVal productionNotes As String, ByVal paragraphStatus As String, ByVal estimatedDuration As String, ByVal paragraphType As String) As String
    Dim chapterNumber As Long, paragraphNumber As Long, paragraphId As String
    chapterNumber = CurrentChapterNumber(ws, paragraphRow)
    If chapterNumber = 0 Then chapterNumber = 1
    paragraphNumber = CountPrefix(ws, Format$(chapterNumber, "00") & "-") + 1
    paragraphId = Format$(chapterNumber, "00") & "-" & Format$(paragraphNumber, "00")
    ws.Rows(paragraphRow & ":" & paragraphRow + 2).Insert Shift:=xlDown
    With ws.Range("A" & paragraphRow)
        .NumberFormat = "@"
        .Value = paragraphId
        .Font.Size = 8
        .Font.Color = RGB(135, 135, 135)
    End With
    ws.Range("B" & paragraphRow).Value = paragraphType
    ws.Range("C" & paragraphRow).Value = paragraphStatus
    ws.Range("D" & paragraphRow).Value = estimatedDuration
    ws.Range("B" & paragraphRow & ":D" & paragraphRow).Interior.Color = RGB(235, 231, 219)
    ws.Range("B" & paragraphRow & ":D" & paragraphRow).Font.Bold = True
    ApplyParagraphPresentation ws, paragraphRow, paragraphType, paragraphStatus
    WriteBlock ws, "B" & paragraphRow + 1 & ":G" & paragraphRow + 1, narration
    WriteBlock ws, "B" & paragraphRow + 2 & ":G" & paragraphRow + 2, productionNotes
    ws.Range("A" & paragraphRow & ":G" & paragraphRow + 2).Borders(xlEdgeLeft).Color = RGB(143, 170, 185)
    ws.Range("A" & paragraphRow & ":G" & paragraphRow + 2).Borders(xlEdgeLeft).Weight = xlThick
    AddParagraphAtRow = paragraphId
End Function

Private Sub WriteSummary(ByVal ws As Worksheet, ByVal address As String, ByVal labelText As String, ByVal valueText As String)
    ws.Range(address).Value = labelText
    ws.Range(address).Font.Bold = True
    ws.Range(address).Font.Color = RGB(73, 100, 119)
    WriteBlock ws, ws.Range(address).Offset(1, 0).Resize(1, 2).Address, valueText
End Sub

Private Sub WriteBand(ByVal ws As Worksheet, ByVal address As String, ByVal valueText As String, ByVal fillColor As Long)
    With ws.Range(address)
        .Merge
        .Value = valueText
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = fillColor
    End With
End Sub

Private Sub WriteBlock(ByVal ws As Worksheet, ByVal address As String, ByVal valueText As String)
    With ws.Range(address)
        .Merge
        .Value = valueText
        .WrapText = True
        .VerticalAlignment = xlTop
        .Interior.Color = RGB(255, 255, 255)
    End With
End Sub

Private Sub WriteChapterField(ByVal ws As Worksheet, ByVal targetRow As Long, ByVal labelText As String, ByVal valueText As String)
    ws.Range("A" & targetRow).Value = labelText
    ws.Range("A" & targetRow).Font.Bold = True
    WriteBlock ws, "B" & targetRow & ":G" & targetRow, valueText
End Sub

Private Sub WriteStoryFlow(ByVal ws As Worksheet)
    ws.Range("A10").Value = "Hook"
    ws.Range("B10").Value = "Context"
    ws.Range("C10").Value = "Tension"
    ws.Range("D10").Value = "Proof"
    ws.Range("E10").Value = "Turning Point"
    ws.Range("F10").Value = "Resolution"
    ws.Range("A10:F10").Interior.Color = RGB(232, 239, 243)
    ws.Range("A10:F10").Font.Bold = True
    ws.Range("A10:F10").HorizontalAlignment = xlCenter
End Sub

Private Sub WriteSidebarGroup(ByVal ws As Worksheet, ByVal startRow As Long, ByVal groupName As String, ByVal helperText As String)
    WriteBand ws, "I" & startRow & ":K" & startRow, groupName, RGB(230, 221, 207)
    ws.Range("I" & startRow & ":K" & startRow).Font.Color = RGB(73, 100, 119)
    WriteBlock ws, "I" & startRow + 1 & ":K" & startRow + 3, helperText
End Sub

Private Function NextContentRow(ByVal ws As Worksheet) As Long
    Dim lastCell As Range
    Set lastCell = ws.Range("A:G").Find(What:="*", After:=ws.Range("A1"), LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then NextContentRow = CONTENT_START_ROW Else NextContentRow = lastCell.Row + 2
End Function

Private Function CountPrefix(ByVal ws As Worksheet, ByVal prefix As String) As Long
    Dim cell As Range, lastRow As Long
    lastRow = NextContentRow(ws)
    For Each cell In ws.Range("A" & CONTENT_START_ROW & ":A" & lastRow).Cells
        If Left$(CStr(cell.Value), Len(prefix)) = prefix Then CountPrefix = CountPrefix + 1
    Next cell
End Function

Private Function CurrentChapterNumber(ByVal ws As Worksheet, ByVal targetRow As Long) As Long
    Dim scanRow As Long, chapterValue As String
    For scanRow = targetRow To CONTENT_START_ROW Step -1
        chapterValue = CStr(ws.Cells(scanRow, "A").Value)
        If Left$(chapterValue, 3) = "CH-" Then
            CurrentChapterNumber = Val(Mid$(chapterValue, 4, 2))
            Exit Function
        End If
    Next scanRow
End Function

Private Function SelectedParagraphInsertRow(ByVal ws As Worksheet, ByVal selectedRow As Long) As Long
    Dim paragraphRow As Long
    paragraphRow = SelectedParagraphRow(ws, selectedRow)
    If paragraphRow > 0 Then SelectedParagraphInsertRow = paragraphRow + 3
End Function
