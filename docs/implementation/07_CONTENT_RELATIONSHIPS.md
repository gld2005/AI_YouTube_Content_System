# Phase 7: Content Relationships

Phase 7 connects the workbook's existing research, packaging, asset, and script records through stable identifiers.

## Delivered behavior

- Evidence can be linked to a selected script paragraph through a Research Link record containing Video ID, Chapter ID, and Paragraph ID.
- New packaging concepts are written to the Packaging Concepts table for a selected video.
- New project asset records can be linked to a specific chapter and paragraph.
- All link operations validate the evidence record before writing a research link and do not infer approval, selection, or final use judgment.

## Validation

`RunPhase7SmokeTest` creates a disposable source, evidence card, video, chapter, paragraph, research link, packaging concept, and project asset link. Expected result:

`PASS: source, evidence, packaging, asset, and paragraph links were created.`

## Requirement traceability

This phase advances AC-082, AC-091, AC-092, AC-120, and AC-121. AI-assisted research parsing, approval-driven evidence generation, and final-paragraph asset synchronization remain open.
