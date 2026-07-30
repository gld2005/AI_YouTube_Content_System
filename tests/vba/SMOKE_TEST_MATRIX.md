# VBA Smoke Test Matrix

Run these tests in a disposable copy after importing the modules into the macro-enabled workbook.

| Test | Procedure | Expected result |
| --- | --- | --- |
| New video | Run `NewVideoScript`; enter a unique ID and title. | A project row, script-index row, current page, and previous page are created. Defaults are Planning, Not Started, P1, today, and V1. |
| Duplicate ID | Run `NewVideoScript`; enter an existing Video ID. | An actionable English message appears and no duplicate rows or sheets are created. |
| Sheet safety | Enter an ID containing invalid sheet-name characters. | The generated sheet names are sanitized and remain unique. |
| Navigation | Select an existing Video ID and run `GoToScriptFromActiveProject`. | The linked current script page opens. |
| Version save | Run `SaveNewVersion`; choose an existing project and confirm. | The current script content overwrites only the paired previous page. No action occurs on workbook open or close. |
| Missing script page | Remove or rename a test script page, then run `SaveNewVersion`. | An actionable English message appears and no content is overwritten. |
| Script workspace | Run `RunPhase4SmokeTest` in a disposable copy. | An English script workspace, one chapter, and a stable-ID three-row paragraph block are created. |
| Script presentation | Run `RunPhase5SmokeTest` in a disposable copy. | Story flow, collapsible production sidebar, four sidebar groups, six paragraph types, and seven statuses are available. |
| Versioning and recycle | Run `RunPhase6SmokeTest` in a disposable copy. | Changed narration is marked before the explicit version replacement, and a deleted paragraph is restored from the recycle store. |
| Security gate | Run `scripts/import_vba_modules.ps1` without AccessVBOM enabled. | The script fails with a clear English message and does not change the Trust Center setting. |
