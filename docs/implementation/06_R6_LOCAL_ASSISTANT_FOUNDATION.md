# R6 Local Assistant Foundation

## Objective

R6 delivers the local-first assistant boundary used by later AI, parsing, and folder-automation stages. It listens on loopback only and does not persist API keys, assistant tokens, or start commands in the workbook or repository.

## Components

| Component | Behavior |
|---|---|
| `assistant/server.py` | Standard-library HTTP service bound to `127.0.0.1` by default. |
| `assistant/config.example.json` | Non-secret configuration example. |
| `assistant/build_executable.ps1` | Reproducible PyInstaller build command for a Windows executable. |
| `src/vba/modAssistantClient.bas` | User-triggered local-assistant start and health-check commands. |
| `tests/assistant/test_server.py` | HTTP schema, authorization, and safe folder-creation coverage. |

## R6 API Contract

The API base is `/v1`, uses UTF-8 JSON, and retains the approved common response shape:

```text
{request_id, operation_id, status, data, warnings, error}
```

`GET /v1/health` is available without a token so Excel can diagnose availability. All other endpoints require `X-Local-Assistant-Token`, whose value is read only from `CONTENT_ASSISTANT_TOKEN` in the assistant process environment.

| Endpoint | R6 behavior |
|---|---|
| `GET /v1/health` | Returns local service version, time, and health state. |
| `GET /v1/capabilities` | Returns the enabled R6 action and future-compatible endpoint list. |
| `POST /v1/folders/create` | Creates confirmed relative template folders below an absolute project root. Absolute child paths and parent traversal are rejected. Existing folders are preserved. |
| `POST /v1/ai/invoke` | Reserved and returns structured `not_enabled`. |
| `POST /v1/sources/parse` | Reserved and returns structured `not_enabled`. |
| `POST /v1/files/parse` | Reserved and returns structured `not_enabled`. |
| `POST /v1/operations/{id}/cancel` | Reserved and returns structured `not_enabled`; R6 has no asynchronous operations. |

The default client timeout is 45 seconds. R6 has no retry loop because it performs no remote AI calls or automatic write retries. Later idempotent asynchronous operations must supply and honor `idempotency_key`; write retries must not occur without one.

## Startup and Trust Model

Set `CONTENT_ASSISTANT_TOKEN` before starting the service. The optional `CONTENT_ASSISTANT_COMMAND` environment variable can contain the user-approved command launched by the Excel `StartLocalAssistant` command. Excel neither stores nor displays either value. `CheckLocalAssistant` uses the named range `nrAssistantBaseUrl`, falling back to `http://127.0.0.1:8765`.

Run the source service with a session-only token:

```powershell
$env:CONTENT_ASSISTANT_TOKEN = "choose-a-random-session-token"
python assistant/server.py
```

Build the Windows executable with the approved external runtime at `E:\CodexTools\pyinstaller-venv`:

```powershell
cd assistant
.\build_executable.ps1
```

The generated executable is stored in `assistant\dist\ContentAssistant.exe`, which is a local project artifact rather than a repository-tracked binary.

## Safety Rules

- Folder creation is a separately confirmed action; R6 does not expose automatic workbook-triggered creation.
- The request requires an absolute project root. Folder template entries cannot leave that root.
- Failure responses identify any paths created before an OS-level failure. Existing folders are never deleted as rollback.
- The same JSON shapes can be retained by a future cloud adapter, where loopback token authentication is replaced by bearer authentication.

## R6 Boundaries

R6 intentionally does not enable AI invocation, source parsing, file parsing, provider credentials, cloud access, or automatic content changes. Those functions remain approval-gated work for later remediation stages.

## Verification Evidence

| Check | Result |
|---|---|
| `tests/assistant/test_server.py` | PASS: health, authorization, folder creation, and path-traversal rejection. |
| `RunR5SmokeTest` in the generated R6 workbook | PASS: existing R5 records workflow was not regressed. |
| R6 workbook generation | PASS: `working/AI_YouTube_Content_System_R6_local_assistant.xlsm` was generated from the macro-free R2 baseline and current `.bas` modules. |
| PyInstaller executable build | PASS: PyInstaller 6.21.0 was installed in `E:\CodexTools\pyinstaller-venv`; `assistant\dist\ContentAssistant.exe` was generated. |
| Executable health check | PASS: the generated executable returned `healthy` over loopback with a temporary session token. |
