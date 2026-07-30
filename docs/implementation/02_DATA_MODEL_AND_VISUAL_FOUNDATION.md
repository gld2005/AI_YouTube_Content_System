# Phase 2: Data Model and Visual Foundation

## Delivered Workbook Foundation

The protected working workbook now contains 22 static worksheets. Per-video current and previous script sheets are intentionally deferred to the VBA creation workflow, because they must be generated from the script template with sanitized, unique sheet names.

## Static Workbook Surfaces

1. `使用说明`
2. `总览`
3. `内容工作台`
4. `选题池`
5. `视频项目`
6. `栏目设置`
7. `结构模板库`
8. `脚本总库`
9. `脚本模板`
10. `研究来源库`
11. `证据卡`
12. `研究关联`
13. `包装实验室`
14. `任务计划`
15. `甘特图`
16. `项目素材`
17. `公共素材库`
18. `标签库`
19. `AI审批中心`
20. `发布复盘`
21. `设置`
22. `迁移审计`

## Data Model Controls

- Every data surface uses an explicitly named Excel table.
- Existing Topic IDs, Video IDs, Source IDs, and Asset IDs were preserved.
- The 38 baseline tasks now have generated stable Task IDs in the documented migration format.
- The original research fact was imported as a review-required evidence candidate, not as automatically approved content.
- The original project status remains available as `Legacy Status`; no project-state transition was confirmed by migration.
- The original thumbnail record was imported as a migrated historical packaging record. No additional concepts were invented.
- Topic scores now use visible formulas referencing editable weights on `设置`.
- Controlled dropdowns are applied to current migrated records. Future VBA actions must extend the same rules when rows are created.

## Visual System

- Base colors: warm off-white and beige.
- Primary accent: low-saturation blue.
- Secondary accent: terracotta for attention and approval-oriented areas.
- Designed sheets hide native gridlines and use explicit warm-gray borders.
- Summary surfaces use compact card-like blocks; operational surfaces use stable, banded tables.
- Chinese body text uses Microsoft YaHei; titles and metrics use Aptos where practical.
- The script template establishes the approved 65/35 writing and production-support layout.

## Verification Evidence

| Check | Result |
| --- | --- |
| Phase 2 workbook worksheets | 22 |
| Named tables | 22 |
| Topics migrated | 2 |
| Video projects migrated | 1 |
| Tasks migrated | 38 |
| Sources migrated | 1 |
| Evidence candidates migrated | 1 |
| Project assets migrated | 3 |
| Publish review records migrated | 1 |
| Formula error scan | 0 matches |
| Rendered sheet review | Completed for all 22 sheets |
| Working workbook SHA-256 | `79F28BDA00843B0257D7467475F8DB8218B0AFAD503987E27E95F28E1E09699C` |

## Acceptance Traceability

This phase advances the following groups while leaving behavior that requires VBA or the local assistant for later phases:

- AC-010 to AC-016: visual language foundation.
- AC-020: dashboard structure and formula-driven stage summary.
- AC-025: content-workspace list surface.
- AC-030 to AC-031: multi-series and structure-template tables.
- AC-070, AC-073, AC-076: visible progress, two-layer state, and threshold fields.
- AC-080 to AC-084: normalized Source Master, Evidence Cards, and Research Links.
- AC-090 to AC-092: Packaging Laboratory schema and historical packaging import.
- AC-100, AC-104 to AC-106: normalized tasks, risk/dependency fields, and Gantt surface.
- AC-120, AC-126 to AC-127: project/shared assets and tag taxonomy tables.
- AC-130 to AC-132: provider configuration layout and no-plaintext-key guidance.
- AC-135 and AC-138: AI Approval Center schema.
- AC-150 to AC-152: visible migration audit, formula scan, and no silent content approval.

## Deferred to Later Phases

- VBA-driven creation, links, formulas that depend on dynamic script pages, versioning, grouping, and folder actions.
- Content Workspace card selection, current-action population, and practical filters.
- Script page generation and paragraph/claim/recycle behavior.
- Gantt scheduling calculations and charts.
- Local assistant, API integrations, AI queue behavior, and executable packaging.
