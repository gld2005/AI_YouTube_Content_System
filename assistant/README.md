# Local Assistant

Set a session-only token and start the service:

```powershell
$env:CONTENT_ASSISTANT_TOKEN = "replace-with-a-random-session-token"
python server.py
```

The service binds to `127.0.0.1:8765` by default. It stores no API keys and accepts no non-loopback connections by default. Build the Windows executable with `build_executable.ps1` after installing PyInstaller in a local build environment.
