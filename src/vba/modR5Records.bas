Attribute VB_Name = "modR5Records"
Option Explicit

Public Function AddOrReuseSource(ByVal sourceType As String, ByVal sourceTitle As String, ByVal sourcePath As String) As String
    Dim sources As ListObject, sourceRow As ListRow
    On Error GoTo Failed
    Set sources = GetTable("SourceMasterTable")
    For Each sourceRow In sources.ListRows
        If StrComp(Trim$(CStr(sourceRow.Range.Cells(1, 8).Value)), Trim$(sourcePath), vbTextCompare) = 0 And Len(Trim$(sourcePath)) > 0 Then
            AddOrReuseSource = CStr(sourceRow.Range.Cells(1, 1).Value)
            Exit Function
        End If
    Next sourceRow
    Set sourceRow = sources.ListRows.Add
    AddOrReuseSource = "SRC-" & Format$(Now, "yyyymmddhhnnss") & "-" & Format$(sources.ListRows.Count, "000")
    SetColumnValue sourceRow, 1, AddOrReuseSource
    SetColumnValue sourceRow, 2, sourceType
    SetColumnValue sourceRow, 3, sourceTitle
    SetColumnValue sourceRow, 7, Date
    SetColumnValue sourceRow, 8, sourcePath
    SetColumnValue sourceRow, 11, "Unverified"
    SetColumnValue sourceRow, 14, "Manual"
    SetColumnValue sourceRow, 15, Now
    Exit Function
Failed:
    AddOrReuseSource = ""
End Function

Public Function AddEvidenceCard(ByVal sourceId As String, ByVal classification As String, ByVal evidenceText As String, ByVal riskLabel As String, ByVal manualJudgment As String) As String
    Dim evidence As ListObject, evidenceRow As ListRow
    On Error GoTo Failed
    If Not IsValidEvidenceClassification(classification) Or Not IsValidManualJudgment(manualJudgment) Then GoTo Failed
    Set evidence = GetTable("EvidenceCardsTable")
    Set evidenceRow = evidence.ListRows.Add
    AddEvidenceCard = "EVD-" & Format$(Now, "yyyymmddhhnnss") & "-" & Format$(evidence.ListRows.Count, "000")
    SetColumnValue evidenceRow, 1, AddEvidenceCard
    SetColumnValue evidenceRow, 2, sourceId
    SetColumnValue evidenceRow, 3, classification
    SetColumnValue evidenceRow, 4, evidenceText
    SetColumnValue evidenceRow, 6, riskLabel
    SetColumnValue evidenceRow, 8, manualJudgment
    SetColumnValue evidenceRow, 12, Now
    Exit Function
Failed:
    AddEvidenceCard = ""
End Function

Public Function CreateDefaultPackagingConcepts(ByVal videoId As String) As Boolean
    Dim packaging As ListObject, row As ListRow, index As Long
    On Error GoTo Failed
    Set packaging = GetTable("PackagingConceptsTable")
    If PackagingCount(videoId) > 0 Then GoTo Failed
    For index = 1 To 5
        Set row = packaging.ListRows.Add
        SetColumnValue row, 1, "PKG-" & videoId & "-" & Format$(index, "00")
        SetColumnValue row, 2, videoId
        SetColumnValue row, 3, index
        SetColumnValue row, 4, "Draft title concept " & index
        SetColumnValue row, 5, "Thumbnail text " & index
        SetColumnValue row, 6, "Define the visual direction."
        SetColumnValue row, 15, "Pending user review"
        SetColumnValue row, 18, "No"
        SetColumnValue row, 19, "No"
    Next index
    RankPackagingConcepts videoId
    CreateDefaultPackagingConcepts = True
    Exit Function
Failed:
    CreateDefaultPackagingConcepts = False
End Function

Public Function RankPackagingConcepts(ByVal videoId As String) As Boolean
    Dim packaging As ListObject, row As ListRow, rankValue As Long, firstScore As Double, secondScore As Double, scoreValue As Double
    On Error GoTo Failed
    Set packaging = GetTable("PackagingConceptsTable")
    For Each row In packaging.ListRows
        If StrComp(CStr(row.Range.Cells(1, 2).Value), videoId, vbTextCompare) = 0 Then
            scoreValue = Val(row.Range.Cells(1, 10).Value) + Val(row.Range.Cells(1, 11).Value) + Val(row.Range.Cells(1, 12).Value)
            SetColumnValue row, 13, scoreValue
            SetColumnValue row, 16, ""
            SetColumnValue row, 17, "No"
            If scoreValue >= firstScore Then
                secondScore = firstScore
                firstScore = scoreValue
            ElseIf scoreValue > secondScore Then
                secondScore = scoreValue
            End If
        End If
    Next row
    rankValue = 0
    For Each row In packaging.ListRows
        If StrComp(CStr(row.Range.Cells(1, 2).Value), videoId, vbTextCompare) = 0 Then
            scoreValue = Val(row.Range.Cells(1, 13).Value)
            If scoreValue = firstScore Or (scoreValue = secondScore And firstScore <> secondScore) Then
                SetColumnValue row, 17, "Yes"
                rankValue = rankValue + 1
                SetColumnValue row, 16, rankValue
            End If
        End If
    Next row
    RankPackagingConcepts = True
    Exit Function
Failed:
    RankPackagingConcepts = False
End Function

Public Function SelectPackagingConcept(ByVal conceptId As String) As Boolean
    Dim packaging As ListObject, row As ListRow, selectedRow As ListRow, videoId As String
    On Error GoTo Failed
    Set packaging = GetTable("PackagingConceptsTable")
    Set selectedRow = FindTableRow(packaging, 1, conceptId)
    If selectedRow Is Nothing Then GoTo Failed
    videoId = CStr(selectedRow.Range.Cells(1, 2).Value)
    For Each row In packaging.ListRows
        If StrComp(CStr(row.Range.Cells(1, 2).Value), videoId, vbTextCompare) = 0 Then SetColumnValue row, 19, "No"
    Next row
    SetColumnValue selectedRow, 18, "Yes"
    SetColumnValue selectedRow, 19, "Yes"
    SelectPackagingConcept = True
    Exit Function
Failed:
    SelectPackagingConcept = False
End Function

Public Function AddProjectAsset(ByVal videoId As String, ByVal assetName As String, ByVal assetType As String, ByVal assetPath As String, ByVal paragraphIds As String) As String
    Dim assets As ListObject, row As ListRow
    On Error GoTo Failed
    Set assets = GetTable("ProjectAssetsTable")
    Set row = assets.ListRows.Add
    AddProjectAsset = "AST-" & Format$(Now, "yyyymmddhhnnss") & "-" & Format$(assets.ListRows.Count, "000")
    SetColumnValue row, 1, AddProjectAsset
    SetColumnValue row, 2, videoId
    SetColumnValue row, 3, assetName
    SetColumnValue row, 4, assetType
    SetColumnValue row, 5, assetPath
    SetColumnValue row, 7, paragraphIds
    SetColumnValue row, 11, "To Find"
    SetColumnValue row, 13, FindSimilarProjectAsset(assetName)
    SetColumnValue row, 15, Now
    Exit Function
Failed:
    AddProjectAsset = ""
End Function

Public Function FindSimilarProjectAsset(ByVal assetName As String) As String
    Dim assets As ListObject, row As ListRow
    On Error GoTo Failed
    Set assets = GetTable("ProjectAssetsTable")
    For Each row In assets.ListRows
        If InStr(1, LCase$(CStr(row.Range.Cells(1, 3).Value)), LCase$(assetName), vbTextCompare) > 0 Or InStr(1, LCase$(assetName), LCase$(CStr(row.Range.Cells(1, 3).Value)), vbTextCompare) > 0 Then
            FindSimilarProjectAsset = CStr(row.Range.Cells(1, 1).Value)
            Exit Function
        End If
    Next row
Failed:
End Function

Public Function PromoteProjectAsset(ByVal assetId As String, ByVal categoryName As String, ByVal tagIds As String) As Boolean
    Dim assets As ListObject, sharedAssetsTable As ListObject, assetRow As ListRow, sharedRow As ListRow
    On Error GoTo Failed
    Set assets = GetTable("ProjectAssetsTable")
    Set assetRow = FindTableRow(assets, 1, assetId)
    If assetRow Is Nothing Then GoTo Failed
    If LCase$(CStr(assetRow.Range.Cells(1, 11).Value)) <> "used" Then GoTo Failed
    Set sharedAssetsTable = GetTable("SharedAssetsTable")
    Set sharedRow = sharedAssetsTable.ListRows.Add
    SetColumnValue sharedRow, 1, assetId
    SetColumnValue sharedRow, 2, assetRow.Range.Cells(1, 3).Value
    SetColumnValue sharedRow, 3, assetRow.Range.Cells(1, 4).Value
    SetColumnValue sharedRow, 4, assetRow.Range.Cells(1, 5).Value
    SetColumnValue sharedRow, 5, categoryName
    SetColumnValue sharedRow, 6, tagIds
    SetColumnValue sharedRow, 8, assetRow.Range.Cells(1, 2).Value
    SetColumnValue sharedRow, 10, "Yes"
    PromoteProjectAsset = True
    Exit Function
Failed:
    PromoteProjectAsset = False
End Function

Public Function AddPendingTag(ByVal tagName As String, ByVal categoryName As String) As String
    Dim tags As ListObject, row As ListRow
    On Error GoTo Failed
    Set tags = GetTable("TagsTable")
    Set row = tags.ListRows.Add
    AddPendingTag = "TAG-" & Format$(Now, "yyyymmddhhnnss") & "-" & Format$(tags.ListRows.Count, "000")
    SetColumnValue row, 1, AddPendingTag
    SetColumnValue row, 2, tagName
    SetColumnValue row, 3, categoryName
    SetColumnValue row, 4, "Pending Organization"
    SetColumnValue row, 6, Date
    SetColumnValue row, 8, "No"
    Exit Function
Failed:
    AddPendingTag = ""
End Function

Private Function PackagingCount(ByVal videoId As String) As Long
    Dim packaging As ListObject, row As ListRow
    Set packaging = GetTable("PackagingConceptsTable")
    For Each row In packaging.ListRows
        If StrComp(CStr(row.Range.Cells(1, 2).Value), videoId, vbTextCompare) = 0 Then PackagingCount = PackagingCount + 1
    Next row
End Function

Private Function IsValidEvidenceClassification(ByVal value As String) As Boolean
    IsValidEvidenceClassification = InStr(1, "Verifiable Fact|Statistic|Quote|Example|Counterpoint|Hypothesis", value, vbTextCompare) > 0
End Function

Private Function IsValidManualJudgment(ByVal value As String) As Boolean
    IsValidManualJudgment = InStr(1, "Use|Do Not Use|Needs Review", value, vbTextCompare) > 0
End Function
