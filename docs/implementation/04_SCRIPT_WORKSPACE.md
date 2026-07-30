# Phase 4: Script Workspace

Phase 4 makes newly created script pages operational rather than template-only. All new developer-facing and user-facing strings introduced by this phase are English-only, following the explicit locale instruction.

## Delivered behavior

- `InitializeScriptWorkspace` creates a 65/35 document and production-sidebar layout on each newly generated current script page.
- `AddChapter` creates an English chapter band with a chapter objective, main claim, subclaims, and transition fields.
- `AddParagraph` inserts the required three-row writing unit after the selected paragraph, with a stable display ID in the `01-01` form, type, status, estimated duration, narration, and production notes.
- `RunPhase4SmokeTest` is non-interactive and validates a fresh project, workspace page, chapter, and paragraph block in a disposable workbook copy.

## Safety and scope

The workspace initializer runs only on the newly copied current script page. It does not alter the reusable template, the source workbook, or an existing project page. The prior-version page remains separate and is still controlled only by the explicit version-save action.

The current phase establishes the visible script structure and stable display IDs. Recycle storage, paragraph difference highlighting, paragraph-type styling, status validation, and asset synchronization remain separate implementation work because they require persistent record and approval models.

## Requirement traceability

This phase advances the structural portions of AC-044, AC-045, AC-047, AC-049, and AC-150. It provides the basis for AC-046, AC-048, AC-050 through AC-052, AC-063 through AC-067, and AC-122 through AC-124, which remain open.

## Validation

Import the VBA modules using `scripts/import_vba_modules.ps1`, then run `RunPhase4SmokeTest` from a disposable macro-enabled workbook copy. The expected result is:

`PASS: English script workspace, chapter, and three-row paragraph block created.`
