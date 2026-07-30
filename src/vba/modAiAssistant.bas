Attribute VB_Name = "modAiAssistant"
Option Explicit

Public Sub PolishCurrentParagraphWithAi()
    RequestCurrentParagraphAiCandidate "polish_paragraph", "Polish Current Paragraph"
End Sub

Public Sub ExpandCurrentParagraphWithAi()
    RequestCurrentParagraphAiCandidate "expand_paragraph", "Expand Current Paragraph"
End Sub

Public Sub CompressCurrentParagraphWithAi()
    RequestCurrentParagraphAiCandidate "compress_paragraph", "Compress Current Paragraph"
End Sub

Public Sub RewriteCurrentParagraphWithAi()
    RequestCurrentParagraphAiCandidate "rewrite_paragraph", "Rewrite Current Paragraph"
End Sub

Public Sub RequestCurrentParagraphAiCandidate(ByVal promptId As String, ByVal featureName As String)
    Dim paragraphRow As Long, videoId As String, chapterId As String, paragraphId As String
    Dim originalText As String, paragraphType As String, paragraphStatus As String, candidateText As String, details As String
    paragraphRow = SelectedParagraphRow(ActiveSheet, ActiveCell.Row)
    If paragraphRow = 0 Then
        ShowUserError "Select a paragraph block before requesting AI assistance.", "No paragraph ID was found above the selected row."
        Exit Sub
    End If
    videoId = VideoIdForScriptSheet(ActiveSheet.Name)
    paragraphId = CStr(ActiveSheet.Cells(paragraphRow, "A").Value)
    chapterId = ChapterIdForRow(ActiveSheet, paragraphRow)
    originalText = CStr(ActiveSheet.Cells(paragraphRow + 1, "B").Value)
    paragraphType = CStr(ActiveSheet.Cells(paragraphRow, "B").Value)
    paragraphStatus = CStr(ActiveSheet.Cells(paragraphRow, "C").Value)
    If Len(Trim$(videoId)) = 0 Or Len(Trim$(originalText)) = 0 Then
        ShowUserError "The selected paragraph is not ready for an AI request.", "Video ID or paragraph text is missing."
        Exit Sub
    End If
    candidateText = RequestAiParagraphCandidate(promptId, originalText, paragraphType, paragraphStatus, "Preserve factual meaning and use supplied evidence only.", details)
    If Len(candidateText) = 0 Then
        ShowUserError "The AI suggestion could not be created. Your script was not changed.", details
        Exit Sub
    End If
    If Len(QueueAiSuggestion(videoId, chapterId, paragraphId, featureName, originalText, candidateText)) = 0 Then
        ShowUserError "The AI suggestion could not be queued for review. Your script was not changed.", "QueueAiSuggestion failed after a valid assistant response."
        Exit Sub
    End If
    MsgBox "The AI suggestion is pending approval. Your script was not changed.", vbInformation, "AI Suggestion"
End Sub

Public Function RequestAiParagraphCandidate(ByVal promptId As String, ByVal paragraphText As String, ByVal paragraphType As String, ByVal paragraphStatus As String, ByVal constraints As String, ByRef details As String) As String
    Dim request As Object, requestBody As String, responseText As String, token As String, quoteCharacter As String
    On Error GoTo Failed
    token = Trim$(Environ$("CONTENT_ASSISTANT_TOKEN"))
    If Len(token) = 0 Then
        details = "CONTENT_ASSISTANT_TOKEN is not available in this Windows session."
        Exit Function
    End If
    quoteCharacter = Chr$(34)
    requestBody = "{" & quoteCharacter & "request_id" & quoteCharacter & ":" & quoteCharacter & Format$(Now, "yyyymmddhhnnss") & quoteCharacter
    requestBody = requestBody & "," & quoteCharacter & "action" & quoteCharacter & ":" & quoteCharacter & "ai.invoke" & quoteCharacter
    requestBody = requestBody & "," & quoteCharacter & "payload" & quoteCharacter & ":{" & quoteCharacter & "provider_id" & quoteCharacter & ":" & quoteCharacter & JsonEscape(AiProviderId()) & quoteCharacter
    requestBody = requestBody & "," & quoteCharacter & "prompt_id" & quoteCharacter & ":" & quoteCharacter & JsonEscape(promptId) & quoteCharacter
    requestBody = requestBody & "," & quoteCharacter & "inputs" & quoteCharacter & ":{" & quoteCharacter & "paragraph" & quoteCharacter & ":" & quoteCharacter & JsonEscape(paragraphText) & quoteCharacter
    requestBody = requestBody & "," & quoteCharacter & "type" & quoteCharacter & ":" & quoteCharacter & JsonEscape(paragraphType) & quoteCharacter
    requestBody = requestBody & "," & quoteCharacter & "status" & quoteCharacter & ":" & quoteCharacter & JsonEscape(paragraphStatus) & quoteCharacter
    requestBody = requestBody & "," & quoteCharacter & "constraints" & quoteCharacter & ":" & quoteCharacter & JsonEscape(constraints) & quoteCharacter & "}}}"
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts 45000, 45000, 45000, 45000
    request.Open "POST", AssistantBaseUrl() & "/v1/ai/invoke", False
    request.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    request.SetRequestHeader "X-Local-Assistant-Token", token
    request.Send requestBody
    responseText = CStr(request.ResponseText)
    If CLng(request.Status) <> 200 Then
        details = "Assistant returned HTTP " & CStr(request.Status) & ". " & JsonStringValue(responseText, "user_message")
        Exit Function
    End If
    If InStr(1, responseText, "pending_approval", vbTextCompare) = 0 Then
        details = "Assistant did not return a pending approval response."
        Exit Function
    End If
    RequestAiParagraphCandidate = JsonStringValue(responseText, "candidate")
    If Len(RequestAiParagraphCandidate) = 0 Then details = "Assistant response did not include a candidate text value."
    Exit Function
Failed:
    details = "AI request failed: " & Err.Description
End Function

Private Function AiProviderId() As String
    AiProviderId = Trim$(Environ$("CONTENT_ASSISTANT_PROVIDER"))
    If Len(AiProviderId) = 0 Then AiProviderId = "openai"
End Function

Private Function JsonEscape(ByVal value As String) As String
    value = Replace(value, Chr$(92), Chr$(92) & Chr$(92))
    value = Replace(value, Chr$(34), Chr$(92) & Chr$(34))
    value = Replace(value, vbCrLf, Chr$(92) & "n")
    value = Replace(value, vbCr, Chr$(92) & "n")
    JsonEscape = Replace(value, vbLf, Chr$(92) & "n")
End Function

Private Function JsonStringValue(ByVal jsonText As String, ByVal keyName As String) As String
    Dim startAt As Long, cursor As Long, result As String, characterText As String, escaped As Boolean, quoteCharacter As String
    quoteCharacter = Chr$(34)
    startAt = InStr(1, jsonText, quoteCharacter & keyName & quoteCharacter, vbTextCompare)
    If startAt = 0 Then Exit Function
    cursor = InStr(startAt, jsonText, ":") + 1
    Do While cursor <= Len(jsonText) And Mid$(jsonText, cursor, 1) = " "
        cursor = cursor + 1
    Loop
    If Mid$(jsonText, cursor, 1) <> quoteCharacter Then Exit Function
    cursor = cursor + 1
    Do While cursor <= Len(jsonText)
        characterText = Mid$(jsonText, cursor, 1)
        If escaped Then
            If characterText = "n" Then result = result & vbLf Else result = result & characterText
            escaped = False
        ElseIf characterText = Chr$(92) Then
            escaped = True
        ElseIf characterText = quoteCharacter Then
            JsonStringValue = result
            Exit Function
        Else
            result = result & characterText
        End If
        cursor = cursor + 1
    Loop
End Function
