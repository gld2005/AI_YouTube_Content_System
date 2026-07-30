# R2 Reproducibility Correction

## Reason for Correction

The R0 backfill independently rebuilt R1 and R2 from their builders. That check found that `scripts/build_r2_workbook.mjs` correctly added the four R2 technical sheets but did not include the Excel-specific finalization that had originally been performed after export.

As a result, a fresh R2 rebuild did not initially contain the required very-hidden technical sheets or the six required named ranges. In addition, three existing assistant configuration names referenced a worksheet name that does not exist in the current Simplified Chinese workbook.

## Applied Correction

`scripts/finalize_r2_workbook.ps1` is now the authoritative R2 finalization step. It:

- Sets `_Lists`, `_ScriptStore`, `_RecycleStore`, and `_MigrationLog` to very hidden.
- Creates or replaces `nrDesignTokens`, `nrParagraphTypeList`, and `nrParagraphStatusList`.
- Creates or replaces `nrProjectRoot`, `nrAssistantBaseUrl`, and `nrAssistantVersion` using the actual settings-sheet name at runtime.
- Avoids hard-coded localized worksheet names in the implementation.

`scripts/build_r2_release.ps1` invokes the existing artifact-tool builder followed by this finalizer, creating a reproducible R2 release workflow.

## Data Preservation

Before applying the correction to the current workbook, its `.xlsx` source was preserved at `backups/R0_20260730/AI_YouTube_Content_System_R2_pre_r2_finalizer.xlsx`.

| Artifact | SHA-256 |
|---|---|
| Pre-finalizer source backup | `3DB450B5D77FC3D00A283F3F5B4BFEA11FD87E07B8FB90FA1AD4D18AAE04AE1E` |
| Corrected R2 source | `10E7BC2C1159C9D26CFEDEC97611D15B392AEE3C3E4A584F9A9AE8426F0747DC` |
| Corrected R2 macro workbook | `34972B41D9541AC0AEDC225E04CB2FD320F8781754B06922D52E6C9E3AF7C89A` |

The correction changes workbook metadata only: technical-sheet visibility and named-range definitions. No business table, project, script, approval, task, or user-visible worksheet content was changed.

## Verification

An independent R1 -> R2 rebuild was finalized and verified. The rebuilt workbook contains all six ranges with valid worksheet references and all four technical sheets are very hidden. The corrected R2 macro workbook also completed the formula-error scan and the Phase 3 through Phase 8 smoke tests on isolated test copies.

## R1 Assessment

R1 requires no adjustment. Its `Design Tokens` sheet, four targeted visual-sheet changes, gridline settings, and frozen rows were reproduced from the R1 builder. The R0 backfill did not change R1 inputs, outputs, or assumptions.
