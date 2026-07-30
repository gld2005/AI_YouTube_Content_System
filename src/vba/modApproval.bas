Attribute VB_Name = "modApproval"
Option Explicit

Public LastApprovalError As String

Public Const AI_APPROVAL_TABLE_NAME As String = "AIApprovalTable"
Public Const PUBLISH_REVIEW_TABLE_NAME As String = "PublishOptimizeTable"

Public Function QueueAiSuggestion(ByVal videoId As String, ByVal chapterId As String, ByVal paragraphId As String, ByVal featureName As String, ByVal originalText As String, ByVal candidateText As String) As String
    Dim approvals As ListObject, newRow As ListRow
    On Error GoTo Failed
    Set approvals = GetTable(AI_APPROVAL_TABLE_NAME)
    Set newRow = approvals.ListRows.Add
    QueueAiSuggestion = "AI-" & Format$(Now, "yyyymmddhhnnss") & "-" & Format$(approvals.ListRows.Count, "000")
    SetColumnValue newRow, 1, QueueAiSuggestion: SetColumnValue newRow, 2, videoId
    SetColumnValue newRow, 3, chapterId
    newRow.Range.Cells(1, 4).NumberFormat = "@"
    newRow.Range.Cells(1, 4).Value = paragraphId
    SetColumnValue newRow, 5, featureName: SetColumnValue newRow, 6, originalText
    SetColumnValue newRow, 7, candidateText: SetColumnValue newRow, 10, "Pending"
    SetColumnValue newRow, 14, "Manual": SetColumnValue newRow, 15, "No external call"
    SetColumnValue newRow, 17, Now: SetColumnValue newRow, 20, "Retain"
    Exit Function
Failed:
    QueueAiSuggestion = ""
End Function

Public Function RejectAiSuggestion(ByVal suggestionId As String, ByVal reason As String) As Boolean
    Dim approvals As ListObject, approvalRow As ListRow
    On Error GoTo Failed
    Set approvals = GetTable(AI_APPROVAL_TABLE_NAME)
    Set approvalRow = FindTableRow(approvals, 1, suggestionId)
    If approvalRow Is Nothing Then GoTo Failed
    SetColumnValue approvalRow, 10, "Rejected": SetColumnValue approvalRow, 12, reason
    SetColumnValue approvalRow, 18, Now: SetColumnValue approvalRow, 20, "Retain rejection reason"
    RejectAiSuggestion = True
    Exit Function
Failed:
    RejectAiSuggestion = False
End Function

Public Function ApproveFullAdoption(ByVal suggestionId As String, ByVal scriptSheet As Worksheet) As Boolean
    Dim approvals As ListObject, approvalRow As ListRow, paragraphId As String, paragraphRow As Long, adoptedText As String
    Dim operation As String
    On Error GoTo Failed
    LastApprovalError = ""
    operation = "locate approvals"
    Set approvals = GetTable(AI_APPROVAL_TABLE_NAME)
    operation = "locate suggestion": Set approvalRow = FindTableRow(approvals, 1, suggestionId)
    If approvalRow Is Nothing Then GoTo Failed
    If CStr(approvalRow.Range.Cells(1, 10).Value) <> "Pending" Then GoTo Failed
    operation = "locate paragraph": paragraphId = CStr(approvalRow.Range.Cells(1, 4).Value)
    paragraphRow = FindParagraphRow(scriptSheet, paragraphId)
    If paragraphRow = 0 Then GoTo Failed
    operation = "write script": adoptedText = CStr(approvalRow.Range.Cells(1, 7).Value)
    scriptSheet.Cells(paragraphRow + 1, "B").Value = adoptedText
    operation = "write approval": SetColumnValue approvalRow, 9, adoptedText: SetColumnValue approvalRow, 10, "Adopted"
    SetColumnValue approvalRow, 11, "Full Adoption": SetColumnValue approvalRow, 18, Now
    operation = "write audit": WriteAuditLine scriptSheet, paragraphId, CStr(approvalRow.Range.Cells(1, 5).Value), "Full Adoption"
    ApproveFullAdoption = True
    Exit Function
Failed:
    LastApprovalError = operation & ": " & Err.Description
    ApproveFullAdoption = False
End Function

Public Function AddPublishReview(ByVal videoId As String, ByVal platformName As String, ByVal finalTitle As String) As Boolean
    Dim reviews As ListObject, newRow As ListRow
    On Error GoTo Failed
    Set reviews = GetTable(PUBLISH_REVIEW_TABLE_NAME)
    Set newRow = reviews.ListRows.Add
    SetColumnValue newRow, 1, videoId: SetColumnValue newRow, 2, platformName
    SetColumnValue newRow, 3, Date: SetColumnValue newRow, 4, finalTitle
    AddPublishReview = True
    Exit Function
Failed:
    AddPublishReview = False
End Function

Private Sub WriteAuditLine(ByVal scriptSheet As Worksheet, ByVal paragraphId As String, ByVal featureName As String, ByVal adoptionMethod As String)
    scriptSheet.Range("I35").Value = "Audit | " & paragraphId & " | " & featureName & " | " & adoptionMethod & " | " & Format$(Now, "yyyy-mm-dd hh:nn")
    scriptSheet.Range("I35").Font.Size = 8
    scriptSheet.Range("I35").Interior.Color = RGB(235, 231, 219)
End Sub
