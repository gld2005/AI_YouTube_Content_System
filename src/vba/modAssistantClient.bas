Attribute VB_Name = "modAssistantClient"
Option Explicit

Private Const ASSISTANT_DEFAULT_URL As String = "http://127.0.0.1:8765"
Private Const ASSISTANT_TIMEOUT_MS As Long = 45000

Public Sub CheckLocalAssistant()
    Dim details As String
    If IsLocalAssistantHealthy(details) Then
        MsgBox "The local assistant is available." & vbCrLf & details, vbInformation, "Local Assistant"
    Else
        ShowUserError "The local assistant is not available. Start it and try again.", details
    End If
End Sub

Public Sub StartLocalAssistant()
    Dim commandText As String
    commandText = Trim$(Environ$("CONTENT_ASSISTANT_COMMAND"))
    If Len(commandText) = 0 Then
        ShowUserError "The local assistant start command is not configured.", "Set CONTENT_ASSISTANT_COMMAND for this Windows session. The workbook does not store assistant commands or tokens."
        Exit Sub
    End If
    On Error GoTo Failed
    CreateObject("WScript.Shell").Run commandText, 0, False
    MsgBox "The local assistant start command was sent. Use Check Local Assistant after it starts.", vbInformation, "Local Assistant"
    Exit Sub
Failed:
    ShowUserError "The local assistant could not be started.", "Assistant startup failed: " & Err.Description
End Sub

Public Function IsLocalAssistantHealthy(ByRef details As String) As Boolean
    Dim request As Object
    Dim responseText As String
    On Error GoTo Failed
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts ASSISTANT_TIMEOUT_MS, ASSISTANT_TIMEOUT_MS, ASSISTANT_TIMEOUT_MS, ASSISTANT_TIMEOUT_MS
    request.Open "GET", AssistantBaseUrl() & "/v1/health", False
    request.Send
    responseText = CStr(request.ResponseText)
    If CLng(request.Status) <> 200 Then
        details = "Health check returned HTTP " & CStr(request.Status) & "."
        Exit Function
    End If
    If InStr(1, responseText, "healthy", vbTextCompare) = 0 Then
        details = "Health check returned an unexpected response."
        Exit Function
    End If
    details = "Health check passed at " & AssistantBaseUrl() & "."
    IsLocalAssistantHealthy = True
    Exit Function
Failed:
    details = "Health check failed: " & Err.Description
End Function

Public Function AssistantBaseUrl() As String
    On Error GoTo UseDefault
    AssistantBaseUrl = Trim$(CStr(ThisWorkbook.Names("nrAssistantBaseUrl").RefersToRange.Value))
    If Len(AssistantBaseUrl) > 0 Then Exit Function
UseDefault:
    AssistantBaseUrl = ASSISTANT_DEFAULT_URL
End Function
