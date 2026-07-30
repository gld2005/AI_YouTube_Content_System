param([string]$PythonPath = "python")

$ErrorActionPreference = "Stop"
& $PythonPath -m PyInstaller --onefile --name ContentAssistant server.py
