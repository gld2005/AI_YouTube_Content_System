# VBA Smoke Test Matrix

Run these tests in a disposable copy after importing the modules into the macro-enabled workbook.

| Test | Procedure | Expected result |
| --- | --- | --- |
| New video | Run `NewVideoScript`; enter a unique ID and title. | A project row, script-index row, current page, and previous page are created. Defaults are Planning, Not Started, P1, today, and V1. |
| Duplicate ID | Run `NewVideoScript`; enter an existing Video ID. | A Chinese actionable message appears and no duplicate rows or sheets are created. |
| Sheet safety | Enter an ID containing invalid sheet-name characters. | The generated sheet names are sanitized and remain unique. |
| Navigation | Select an existing Video ID and run `GoToScriptFromActiveProject`. | The linked current script page opens. |
| Version save | Run `SaveNewVersion`; choose an existing project and confirm. | The current script content overwrites only the paired previous page. No action occurs on workbook open or close. |
| Missing script page | Remove or rename a test script page, then run `SaveNewVersion`. | A Chinese actionable message appears and no content is overwritten. |
| Security gate | Run `scripts/import_vba_modules.ps1` without AccessVBOM enabled. | The script fails with a clear English message and does not change the Trust Center setting. |
