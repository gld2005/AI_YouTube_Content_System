# R5 Assets, Packaging, and Research Records

## Objective

R5 adds deterministic record workflows for reusable research sources, evidence, packaging concepts, project assets, shared assets, and tag governance. AI parsing and semantic matching remain in R7 and R8; R5 does not make an AI-generated record authoritative.

## Components

| Component | Behavior |
|---|---|
| `src/vba/modR5Records.bas` | Adds source reuse, evidence cards, five default packaging concepts, scoring and selection, project assets, text-based similarity candidates, shared-asset promotion, and pending tags. |
| `src/vba/modSmokeTests.bas` | Adds `RunR5SmokeTest`. |
| `working/AI_YouTube_Content_System_R5_final2.xlsm` | Verified R5 macro workbook. |

## R5 Rules

- A source URL or local path is reused rather than silently duplicated.
- One source can hold multiple evidence cards. Evidence classification and manual judgment inputs are constrained to their accepted values.
- A new video receives five packaging concepts. The user explicitly selects one concept; selection overrides rank.
- Packaging scoring is deterministic from content match, differentiation, and credibility scores. The two highest-scoring concepts are flagged.
- Project assets begin in `To Find`, store production-oriented fields, and expose a text-matched similar-asset candidate.
- Only a `Used` project asset can be promoted to the shared library.
- New custom tags begin as `Pending Organization` and remain inactive until a later approval workflow.

## Data Protection

The R4 workbook was preserved before implementation at `backups/R5_20260731/AI_YouTube_Content_System_R4_pre_R5.xlsm` with SHA-256 `2AA8D415D31422DA7D52A18EBC3DA4226BB12DF44D5B3163561D4C912C2E1291`.

R5 adds modules only. It does not remove, reorder, or overwrite existing model-table data.

## Verification Evidence

| Check | Result |
|---|---|
| VBA compilation in Windows Excel | Passed |
| `RunR5SmokeTest` | PASS |
| Source reuse fixture | PASS |
| Multiple-evidence fixture | PASS |
| Five-concept packaging fixture | PASS |
| Direct concept selection fixture | PASS |
| Used-asset promotion fixture | PASS |
| Pending-tag fixture | PASS |
| Phase 3 through Phase 8 regressions | PASS |
| Formula-error scan | 0 errors |

## R5 Exit Condition

R5 is complete for deterministic record workflows. AI semantic asset matching, AI packaging recommendations, provider configuration, approval-gated generated content, and automated source parsing remain assigned to R6 through R8.
