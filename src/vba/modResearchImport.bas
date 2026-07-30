Attribute VB_Name = "modResearchImport"
Option Explicit

Public Sub ImportResearchUrl()
    Dim sourceUrl As String, sourceTitle As String, request As Object, responseText As String, token As String, requestBody As String
    sourceUrl = Trim$(InputBox("Enter a public source URL:", "Import Research Source"))
    If Len(sourceUrl) = 0 Then Exit Sub
    sourceTitle = Trim$(InputBox("Enter a source title for review:", "Import Research Source"))
    If Len(sourceTitle) = 0 Then Exit Sub
    token = Trim$(Environ$("CONTENT_ASSISTANT_TOKEN"))
    If Len(token) = 0 Then
        ShowUserError "The local assistant token is not available. Your source was not created.", "CONTENT_ASSISTANT_TOKEN is unavailable."
        Exit Sub
    End If
    On Error GoTo Failed
    requestBody = "{" & Chr$(34) & "payload" & Chr$(34) & ":{" & Chr$(34) & "url" & Chr$(34) & ":" & Chr$(34) & Replace(sourceUrl, Chr$(34), Chr$(92) & Chr$(34)) & Chr$(34) & "}}"
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts 45000, 45000, 45000, 45000
    request.Open "POST", AssistantBaseUrl() & "/v1/sources/parse", False
    request.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    request.SetRequestHeader "X-Local-Assistant-Token", token
    request.Send requestBody
    responseText = CStr(request.ResponseText)
    If CLng(request.Status) <> 200 Then
        ShowUserError "The source could not be parsed. Use the recovery instructions and add it manually if needed.", responseText
        Exit Sub
    End If
    If Len(AddOrReuseSource("Web", sourceTitle, sourceUrl)) = 0 Then
        ShowUserError "The parsed source could not be recorded.", "AddOrReuseSource failed after a successful parse."
        Exit Sub
    End If
    MsgBox "Source metadata was recorded. Review extracted content before creating evidence.", vbInformation, "Research Import"
    Exit Sub
Failed:
    ShowUserError "The source import failed. Your research records were not changed.", Err.Description
End Sub
