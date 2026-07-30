# R3 Script Workspace and VBA UX

## Objective

R3 strengthens the interactive script workspace on top of the R2 workbook without changing project, script, approval, task, or research records.

## Affected Components

| Component | Change |
|---|---|
| `src/vba/modScriptWorkspace.bas` | Adds six visible rounded action buttons in the production sidebar. |
| `src/vba/modScriptRecycle.bas` | Requires confirmation before user-triggered paragraph or chapter deletion. |
| `src/vba/modVersioning.bas` | Replaces the previous-version sheet using a staged copy/rename sequence rather than deleting it before the replacement copy exists. |
| `src/vba/modSmokeTests.bas` | Uses explicit non-interactive test-only recycle helpers so destructive-action confirmation remains a user-facing safeguard. |
| `working/AI_YouTube_Content_System_R3.xlsm` | R3 macro workbook built from the corrected R2 source. |

## Interaction Behavior

The R3 script workspace places the document area in columns `A:G`, a gutter in `H`, and the production sidebar in `I:K`. Its six sidebar buttons are `Add Chapter`, `Add Paragraph`, `Save Version`, `Delete Paragraph`, `Delete Chapter`, and `Restore Item`.

The two deletion actions display a Yes/No warning before any content is moved to the recycle area. Choosing No performs no write. The existing explicit version-save confirmation remains in place; R3 also ensures the existing previous-version sheet remains intact until a replacement copy is ready to be published.

## Migration and Compatibility

The source R2 macro workbook was preserved at `backups/R3_20260731/AI_YouTube_Content_System_R2_pre_R3.xlsm` with SHA-256 `34972B41D9541AC0AEDC225E04CB2FD320F8781754B06922D52E6C9E3AF7C89A`.

No table schema, source-of-truth record, visible model sheet, formula, or named range was changed. R3 targets Windows desktop Excel with macros enabled.

## Verification Evidence

| Check | Result |
|---|---|
| Clean R3 macro build | Passed |
| Formula-error scan | 0 errors |
| Script action shapes | 6 created with the expected macro bindings |
| Phase 3 smoke test | PASS |
| Phase 4 smoke test | PASS |
| Phase 5 smoke test | PASS |
| Phase 6 smoke test | PASS |
| Phase 7 smoke test | PASS |
| Phase 8 smoke test | PASS |

## Acceptance Scope

R3 provides and verifies the interactive script actions, explicit destructive-action warning, paragraph/recycle workflow, current/previous version workflow, script navigation, and approval-gated adoption already implemented by the VBA modules. Remaining schedule, asset, assistant, provider, and source-parser capabilities remain assigned to R4 through R8.
