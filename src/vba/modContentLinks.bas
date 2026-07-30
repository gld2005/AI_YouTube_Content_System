Attribute VB_Name = "modContentLinks"
Option Explicit

Public Const SOURCE_TABLE_NAME As String = "SourceMasterTable"
Public Const EVIDENCE_TABLE_NAME As String = "EvidenceCardsTable"
Public Const RESEARCH_LINK_TABLE_NAME As String = "ResearchLinksTable"
Public Const PACKAGING_TABLE_NAME As String = "PackagingConceptsTable"
Public Const PROJECT_ASSET_TABLE_NAME As String = "ProjectAssetsTable"

Public Sub LinkEvidenceToSelectedParagraph()
    Dim evidenceId As String, paragraphRow As Long, paragraphId As String, videoId As String, chapterId As String
    paragraphRow = SelectedParagraphRow(ActiveSheet, ActiveCell.Row)
    If paragraphRow = 0 Then
        ShowUserError "Select a paragraph block before linking evidence.", "Evidence linking requires a selected paragraph block."
        Exit Sub
    End If
    evidenceId = Trim$(InputBox("Enter an existing Evidence ID:", "Link Evidence"))
    If Len(evidenceId) = 0 Then Exit Sub
    paragraphId = CStr(ActiveSheet.Cells(paragraphRow, "A").Value)
    videoId = VideoIdForScriptSheet(ActiveSheet.Name)
    chapterId = ChapterIdForRow(ActiveSheet, paragraphRow)
    If Not AddResearchLink(videoId, evidenceId, chapterId, paragraphId, "Script evidence") Then
        ShowUserError "The evidence link could not be created.", "Evidence link creation failed for " & evidenceId
    End If
End Sub

Public Function AddResearchLink(ByVal videoId As String, ByVal evidenceId As String, ByVal chapterId As String, ByVal paragraphId As String, ByVal linkPurpose As String) As Boolean
    Dim evidence As ListObject, links As ListObject, evidenceRow As ListRow, newRow As ListRow
    On Error GoTo Failed
    Set evidence = GetTable(EVIDENCE_TABLE_NAME)
    Set evidenceRow = FindTableRow(evidence, 1, evidenceId)
    If evidenceRow Is Nothing Then GoTo Failed
    Set links = GetTable(RESEARCH_LINK_TABLE_NAME)
    Set newRow = links.ListRows.Add
    SetColumnValue newRow, 1, "RL-" & Format$(Now, "yyyymmddhhnnss") & "-" & Format$(links.ListRows.Count, "000")
    SetColumnValue newRow, 2, CStr(evidenceRow.Range.Cells(1, 2).Value)
    SetColumnValue newRow, 3, evidenceId
    SetColumnValue newRow, 4, videoId
    SetColumnValue newRow, 6, chapterId
    SetColumnValue newRow, 8, paragraphId
    SetColumnValue newRow, 9, linkPurpose
    SetColumnValue newRow, 10, "Yes"
    AddResearchLink = True
    Exit Function
Failed:
    AddResearchLink = False
End Function

Public Function AddPackagingConcept(ByVal videoId As String, ByVal titleText As String, ByVal thumbnailText As String) As Boolean
    Dim concepts As ListObject, newRow As ListRow
    On Error GoTo Failed
    Set concepts = GetTable(PACKAGING_TABLE_NAME)
    Set newRow = concepts.ListRows.Add
    SetColumnValue newRow, 1, "PC-" & Format$(Now, "yyyymmddhhnnss") & "-" & Format$(concepts.ListRows.Count, "000")
    SetColumnValue newRow, 2, videoId: SetColumnValue newRow, 3, concepts.ListRows.Count
    SetColumnValue newRow, 4, titleText: SetColumnValue newRow, 5, thumbnailText
    SetColumnValue newRow, 15, "Pending Review": SetColumnValue newRow, 19, "No"
    AddPackagingConcept = True
    Exit Function
Failed:
    AddPackagingConcept = False
End Function

Public Function AddProjectAssetLink(ByVal videoId As String, ByVal chapterId As String, ByVal paragraphId As String, ByVal assetName As String) As Boolean
    Dim assets As ListObject, newRow As ListRow
    On Error GoTo Failed
    Set assets = GetTable(PROJECT_ASSET_TABLE_NAME)
    Set newRow = assets.ListRows.Add
    SetColumnValue newRow, 1, "AS-" & Format$(Now, "yyyymmddhhnnss") & "-" & Format$(assets.ListRows.Count, "000")
    SetColumnValue newRow, 2, videoId: SetColumnValue newRow, 3, assetName
    SetColumnValue newRow, 4, "Reference": SetColumnValue newRow, 6, chapterId
    SetColumnValue newRow, 7, paragraphId: SetColumnValue newRow, 11, "Planned"
    SetColumnValue newRow, 15, Now
    AddProjectAssetLink = True
    Exit Function
Failed:
    AddProjectAssetLink = False
End Function

Public Function VideoIdForScriptSheet(ByVal scriptSheetName As String) As String
    Dim projects As ListObject, row As ListRow
    On Error GoTo Failed
    Set projects = GetTable(PROJECT_TABLE_NAME)
    For Each row In projects.ListRows
        If StrComp(CStr(row.Range.Cells(1, 28).Value), scriptSheetName, vbTextCompare) = 0 Then
            VideoIdForScriptSheet = CStr(row.Range.Cells(1, 1).Value)
            Exit Function
        End If
    Next row
Failed:
End Function

Public Function ChapterIdForRow(ByVal ws As Worksheet, ByVal targetRow As Long) As String
    Dim rowIndex As Long
    For rowIndex = targetRow To 12 Step -1
        If Left$(CStr(ws.Cells(rowIndex, "A").Value), 3) = "CH-" Then
            ChapterIdForRow = Left$(CStr(ws.Cells(rowIndex, "A").Value), 5)
            Exit Function
        End If
    Next rowIndex
End Function
