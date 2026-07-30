# Local Assistant

Set a session-only token and start the service:

```powershell
$env:CONTENT_ASSISTANT_TOKEN = "replace-with-a-random-session-token"
python server.py
```

The service binds to `127.0.0.1:8765` by default. It stores no API keys and accepts no non-loopback connections by default. Build the Windows executable with `build_executable.ps1`; its default Python runtime is installed outside the repository at `E:\CodexTools\pyinstaller-venv`. The executable output is `dist\ContentAssistant.exe` in this project directory.

For an AI request, set `CONTENT_ASSISTANT_PROVIDER` to a provider ID in `providers.json` and set that provider's API-key environment variable. The selected provider's model can be overridden with its documented model environment variable. Never put API keys in the workbook, command line, or a committed file.
