# R4 Status and Schedule Foundation

## Objective

R4 introduces the first deterministic VBA foundation for project progress, status suggestions, task progress, and the Gantt feed. It does not automatically transition project status or working stage.

## Affected Components

| Component | Behavior |
|---|---|
| `src/vba/modProjectStatus.bas` | Calculates structure, writing, recording, asset, and editing progress; applies the 20/35/15/15/15 total weighting; calculates remaining days and risk; displays an advisory-only stage suggestion. |
| `src/vba/modTasks.bas` | Calculates equal-weight fallback child-task progress, derives task risk, and rebuilds `GanttTable` from main tasks and selected key child tasks. |
| `working/AI_YouTube_Content_System_R4.xlsm` | R4 macro workbook built from the protected R3 source. |

## Controls and Rules

- Paragraph writing progress uses estimated duration as its weight where a valid `mm:ss` duration exists; invalid or missing duration uses a safe weight of one.
- Structure progress requires populated chapter planning fields and an explicit `Confirmed` value in the chapter confirmation cell.
- Total progress is calculated as Structure 20%, Writing 35%, Recording 15%, Assets 15%, and Edit 15%.
- Project status and working stage are never changed by the suggestion function. It displays a proposed stage only, leaving confirmation to the user.
- Gantt refresh includes task rows marked as a main task or key child task. It copies the planned start, due date, current progress, and derived risk into `GanttTable`.
- Task risk is `Blocked` when a blocker reason exists, otherwise `Overdue`, `Due Soon`, or `Normal` based on due date and completion state.

## Migration and Compatibility

The R3 macro workbook was preserved before R4 at `backups/R4_20260731/AI_YouTube_Content_System_R3_pre_R4.xlsm` with SHA-256 `F17E09E1BEE78219D9B1729DE07D6B587A0F12B120C36AB5532D2EF5D42679BD`.

R4 adds VBA modules only. It does not change model tables, existing records, formulas, names, or visible worksheets.

## Verification Evidence

| Check | Result |
|---|---|
| R4 macro workbook build | Passed |
| Phase 3 through Phase 8 regression suite | PASS |
| `RefreshGanttForVideo` isolated invocation | PASS |
| Existing workbook formula-error scan | 0 errors before R4 macro packaging |
| Automated non-interactive status-refresh acceptance | Deferred to Windows Excel manual acceptance |

## Remaining R4 Acceptance Work

The progress and Gantt functions are implemented, but the completed system still requires Windows Excel manual acceptance for the project-status suggestion path and visual Gantt rendering. Reverse-planning UI, series threshold overrides, detailed dependency impact display, and a visual stage timeline remain planned R4 completion items and require a further approved R4 increment before they can be reported as implemented.
