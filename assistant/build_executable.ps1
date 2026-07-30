param([string]$PythonPath = "E:\CodexTools\pyinstaller-venv\Scripts\python.exe")

$ErrorActionPreference = "Stop"
if (-not (Test-Path -LiteralPath $PythonPath)) {
    throw "PyInstaller Python was not found at: $PythonPath"
}
& $PythonPath -m PyInstaller --onefile --name ContentAssistant --add-data "prompts;prompts" --add-data "providers.json;." server.py
