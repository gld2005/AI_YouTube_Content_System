# Baseline Migration Plan

## Purpose

This plan preserves data from `reference/AI_YouTube_Content_Pipeline_Baseline.xlsx` while moving the system to the approved content-first architecture. The reference workbook is immutable. All structural work starts from `working/AI_YouTube_Content_System_Working.xlsx` after a verified backup is made.

## Source Integrity

| Item | Value |
| --- | --- |
| Source workbook | `reference/AI_YouTube_Content_Pipeline_Baseline.xlsx` |
| SHA-256 | `B92E95248398294C1D847D2B0450E1CF01185F59DE5E19888FFCA1AAD65649A1` |
| Logical source sheets | 9 |
| Formula error scan | No matches for `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, or `#N/A` |

## Migration Rules

1. Preserve existing record values and source identifiers before adding new fields.
2. Never infer an approved script, evidence judgment, series, packaging concept, or project-state transition from incomplete baseline data.
3. Retain unmapped legacy information in a visible `Legacy Notes` field or a dedicated import-audit table.
4. Create stable IDs only where the baseline does not already provide one; save every generated-ID crosswalk.
5. Create a timestamped backup of the working workbook before every structural migration.
6. Run the baseline assertions before and after migration. A failed assertion blocks export.

## Sheet and Field Mapping

| Baseline surface | Target surface | Preservation and transformation rule |
| --- | --- | --- |
| `使用说明` | `使用说明` | Preserve instructional content as a refreshed user guide. Keep legacy guidance in an import note until the replacement guide is approved. |
| `总览` | `总览`, `内容工作台`, `_Config` | Preserve the selected Video ID as a configuration value. Rebuild summary values, action list, and chart from normalized project and task records; do not treat dashboard cells as authoritative duplicates. |
| `选题池` | `选题池` | Preserve Topic ID, title, pillar, all seven scores, links, decision, format, duration, dates, and notes. Recalculate total score from configurable weights and retain any source score in `Legacy Notes` if it differs. |
| `视频项目` | `视频项目` | Preserve Video ID, title, topic, pillar, format, duration, audience, promise, score, status, priority, owner, dates, folder link, last update, and notes. Add missing target-model fields as blank or documented defaults. Existing records remain `Unclassified` until a user assigns a series. |
| `任务计划` | `任务计划` | Import each baseline row as a Project-layer main task unless a later user decision establishes a parent/child relationship. Generate stable Task IDs and retain the source row in the migration crosswalk. Preserve stage, owner, dates, dependencies, deliverable, notes, and legacy completion value. |
| `研究库` | `研究来源库`, `证据卡`, `研究关联` | Create Source Master records from source-level metadata. Convert a populated core-fact field into an evidence-card candidate with `Needs More Evidence` and `Use with Careful Wording`; no AI or migration process may mark it authoritative. Preserve quotes, bias, and intended use in the source or evidence notes. |
| `素材清单` | `项目素材`, `公共素材库` | Preserve Asset ID, video links, type, description, purpose, file/URL, status, owner, dates, filename, placement, and notes. Move legacy source-method and rights fields into `Production Notes` until a dedicated compatibility column is no longer needed. Do not promote any asset to the shared library automatically. |
| `发布复盘` | `发布复盘`, `包装实验室` | Preserve all publishing and performance fields. Create at most one historical packaging record when a baseline title or thumbnail version exists; mark it as migrated history, not an AI recommendation. Do not manufacture the other four concepts. |
| `设置` | `设置`, `_Lists` | Preserve topic-scoring weights and legacy task status conventions. Add the approved controlled lists, thresholds, provider settings, and security fields separately. |

## Identifier Strategy

| Entity | Migration rule |
| --- | --- |
| Topic, Video, Source, Asset | Preserve existing IDs exactly. |
| Task | Generate `TASK-{VideoID}-{StageSequence}-{Ordinal}` and store the source sheet and row number in the crosswalk. |
| Evidence | Generate `EVD-{SourceID}-01` only for a populated baseline evidence candidate. |
| Packaging Concept | Generate `PKG-{VideoID}-01` only for a historical selected title or thumbnail record. |
| Series, Structure Template, Chapter, Claim, Paragraph, Research Link, Tag, AI Suggestion, Version, Deleted Item | Do not fabricate records. Create only through approved workbook actions or explicit migration decisions. |

## Required Import Audit Tables

- `_MigrationCrosswalk`: source workbook, sheet, row, source identifier, target identifier, migration timestamp, migration status, and notes.
- `_MigrationExceptions`: records blocked by invalid IDs, ambiguous links, invalid dates, or unsupported values.
- `_MigrationSummary`: input/output counts, preserved counts, warning counts, and source/working hashes.

## Acceptance Coverage

This phase supports AC-001 to AC-005, AC-080, AC-082, AC-120, AC-143, AC-150 to AC-154. It establishes preservation controls required before the broader feature work mapped in `docs/07_REQUIREMENTS_TRACEABILITY.md`.

## Exit Criteria

1. The source hash matches the approved baseline manifest.
2. The working-copy hash matches the source before structural migration begins.
3. Mapping rules and exception behavior are reviewed before writing any migrated data.
4. The migration test fixture is committed with this plan.
