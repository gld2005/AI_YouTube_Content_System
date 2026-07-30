# Phase 3: VBA Core

## Delivered Source Modules

- `modCore.bas`: named-table lookup, header resolution, stable sheet-name sanitization, safe table writes, and English technical logging.
- `modNavigation.bas`: Dashboard and Content Workspace navigation plus script navigation from a selected project ID.
- `modNewVideo.bas`: duplicate-ID validation, project and script-index record creation, documented defaults, and paired current/previous script sheets copied from the template.
- `modVersioning.bas`: an explicit, confirmed current-to-previous script copy action. It never runs during workbook open or close.

## Security Gate

The local Excel installation exposes COM automation, but the Trust Center `AccessVBOM` setting is not enabled. This is a user-controlled global security setting. The project does not alter it.

To import and compile the VBA modules, the workbook owner must explicitly enable **Trust access to the VBA project object model** in Excel Trust Center, then run:

```powershell
.\scripts\import_vba_modules.ps1
```

The import script creates `working/AI_YouTube_Content_System_Working.xlsm` and imports all `.bas` modules. It fails closed when the setting is absent.

## Deferred Behavior

The script page currently provides the required visual template only. Dynamic chapter/paragraph operations, stable paragraph IDs, recycle storage, restoration, and paragraph-level difference highlighting require the Phase 4 script-record model and will be implemented against that structure. This avoids fabricating or deleting content before the record model exists.

## Acceptance Traceability

This phase advances AC-040 to AC-043, AC-060 to AC-062, AC-150, and AC-153 through source-controlled VBA behavior. AC-063 to AC-067 and the remaining script-workspace behavior remain open for Phase 4.
