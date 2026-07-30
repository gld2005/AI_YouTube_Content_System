# R0 Remediation Baseline

## Purpose

R0 was completed after R1 and R2 had already been applied. It establishes an immutable pre-remediation reference point for the current R2 workbook; it does not change business data, workbook content, VBA source, or the R2 deliverable.

## Protected Backup

| Item | Location | SHA-256 |
|---|---|---|
| Current R2 macro workbook | `working/AI_YouTube_Content_System_R2.xlsm` | `DD3B040EDA2D5F30A3F8D2F7AF6AABB7FD86681767A12B10E181D39829C85606` |
| Immutable R0 backup | `backups/R0_20260730/AI_YouTube_Content_System_R2_pre_remediation.xlsm` | `DD3B040EDA2D5F30A3F8D2F7AF6AABB7FD86681767A12B10E181D39829C85606` |

The matching hashes verify that the preserved backup is byte-identical to the R2 macro workbook at the time R0 was executed.

## Workbook Structure Snapshot

The R2 workbook contains 27 worksheets: 23 visible worksheets and four very-hidden technical worksheets. The source workbook UI remains Simplified Chinese; the technical `Design Tokens` sheet is English by design.

| Index | Worksheet | Visibility | Used range |
|---:|---|---|---|
| 1 | 使用说明 | Visible | A1:J9 |
| 2 | 总览 | Visible | A1:N32 |
| 3 | 内容工作台 | Visible | A1:N32 |
| 4 | 选题池 | Visible | A1:S5 |
| 5 | 视频项目 | Visible | A1:AH4 |
| 6 | 栏目设置 | Visible | A1:S4 |
| 7 | 结构模板库 | Visible | A1:M3 |
| 8 | 脚本总库 | Visible | A1:N4 |
| 9 | 脚本模板 | Visible | A1:R32 |
| 10 | 研究来源库 | Visible | A1:P4 |
| 11 | 证据卡 | Visible | A1:L4 |
| 12 | 研究关联 | Visible | A1:J4 |
| 13 | 包装实验室 | Visible | A1:AA4 |
| 14 | 任务计划 | Visible | A1:AF41 |
| 15 | 甘特图 | Visible | A1:N32 |
| 16 | 项目素材 | Visible | A1:P6 |
| 17 | 公共素材库 | Visible | A1:J3 |
| 18 | 标签库 | Visible | A1:H3 |
| 19 | AI审批中心 | Visible | A1:S3 |
| 20 | 发布复盘 | Visible | A1:Y4 |
| 21 | 设置 | Visible | A1:L10 |
| 22 | 迁移审计 | Visible | A1:G4 |
| 23 | Design Tokens | Visible | A1:C10 |
| 24 | _Lists | Very hidden | A1:D8 |
| 25 | _ScriptStore | Very hidden | A1:J2 |
| 26 | _RecycleStore | Very hidden | A1:J2 |
| 27 | _MigrationLog | Very hidden | A1:J2 |

The table inventory is preserved in the R0 snapshot: `TopicsTable` (2 rows), `VideoProjectsTable` (1), `SeriesTable` (1), `StructureTemplatesTable` (1), `ScriptIndexTable` (1), `SourceMasterTable` (1), `EvidenceCardsTable` (1), `ResearchLinksTable` (1), `PackagingConceptsTable` (1), `TasksTable` (38), `GanttTable` (1), `ProjectAssetsTable` (3), `SharedAssetsTable` (1), `TagsTable` (1), `AIApprovalTable` (1), `PublishOptimizeTable` (1), and `MigrationAuditTable` (1).

Named ranges present at baseline: `nrAssistantBaseUrl`, `nrAssistantVersion`, `nrDesignTokens`, `nrParagraphStatusList`, `nrParagraphTypeList`, and `nrProjectRoot`. The workbook also retains Excel compatibility names `_xlfn.COUNTIFS` and `_xlfn.IFERROR`.

## VBA Source Manifest

The following exported source modules are the baseline for future remediation comparison:

| Module | SHA-256 |
|---|---|
| `modApproval.bas` | `67E121A01FA1E00BE8CF54F47D511519D666EF2E1387998258FE7477B40EE5A1` |
| `modContentLinks.bas` | `21E166BE52F3777B3D9B009D6A447733D219CF3135CEA760FDA962B64B3254BF` |
| `modCore.bas` | `A2E81C07B655DEB1925F26F394D4F71F20F0BFD7B03C03E940F3B9B1A37F6214` |
| `modEnglishUiMigration.bas` | `02952BF99ECBD50BB19C1A12C035A48663C0700DECF36C5D46CB21A2D065F345` |
| `modNavigation.bas` | `344B97B0C342E3D5EBDF0C4DF96EEC9C8BB16DB40B7E77AEAA1AF6B1F9B805DD` |
| `modNewVideo.bas` | `AA8C91B43D1C9B1100F43198FC9154DFC1DAAA242114529D21C02AA16D2B9659` |
| `modScriptControls.bas` | `894E9AC425BBAD4D40DF7C45C2C4E581842E54660D7AF74C428679E6086DFE65` |
| `modScriptRecycle.bas` | `923ABF12CE02E5836F6DF4CC5631D18170A4F02EF1C45D95A99597B6DCF783DD` |
| `modScriptWorkspace.bas` | `3319EC77FC38BD1D21AB5D919648936E4A714620A9F22BF89D806346BD7B6ED4` |
| `modSmokeTests.bas` | `3DF345AB13D873FB1803DBF047386D19BFD7751FD884CCA8F0A8C209B3D39F95` |
| `modVersioning.bas` | `CDD6BF88E4F619C987B099769A796F14DC343D4EF082DD644A62B7CF4DF089C1` |

## Clean Build and Regression Evidence

A clean macro workbook was created from the R2 `.xlsx` source by importing the exported modules with `scripts/import_vba_modules.ps1`. It is intentionally separate from R2.

| Check | Result |
|---|---|
| Clean-build artifact | `working/AI_YouTube_Content_System_R0_cleanbuild.xlsm` |
| Clean-build SHA-256 | `98D9F4609DF5A13952163892C0EA8C2C165E728082630B6A719446DCC19368CA` |
| Formula error scan | 0 cells with formula errors |
| `RunPhase3SmokeTest` | PASS |
| `RunPhase4SmokeTest` | PASS |
| `RunPhase5SmokeTest` | PASS |
| `RunPhase6SmokeTest` | PASS |
| `RunPhase7SmokeTest` | PASS |
| `RunPhase8SmokeTest` | PASS |

Each smoke test was executed against a separate disposable copy of the clean-build workbook. This avoids smoke-test records changing the R2 deliverable or its protected backup.

## R0 Exit Condition

R0 is complete: the protected backup, inventory, source manifest, clean build, and regression evidence are available. No remediation work beyond the already-approved R1 and R2 changes is authorized by this record. Proceed only after explicit approval of the next phase.
