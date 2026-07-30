# Phase 6: Versioning and Recycle

Phase 6 adds explicit script comparison and reversible deletion behavior. All source strings and technical messages remain English-only.

## Delivered behavior

- `SaveVersionForVideo` compares the current script with the retained previous script before replacing that previous version.
- A changed or newly added paragraph is marked with light terracotta in its narration and production-note cells. Unchanged paragraph content returns to the neutral warm off-white background.
- The prior script is replaced only after the explicit version action; opening or closing the workbook does not create a version.
- `DeleteSelectedParagraph` and `DeleteSelectedChapter` store a formatted snapshot before removing the selected block.
- `RestoreSelectedRecycleItem` and `RestoreRecycleItem` restore a stored block without changing its stable paragraph ID.
- Each script page has a collapsed recycle list. Snapshot metadata and copied source blocks are held in a very-hidden workbook sheet, not in cells visible during normal editing.
- The recycle store enforces a maximum of 20 active items per script and does not clear items when a new version is saved.

## Safety model

The version action still requires the existing user confirmation in `SaveNewVersion`. It retains exactly a current and previous script page. Recycle storage is separate from versioning, so saving a version never deletes recoverable items.

## Validation

`RunPhase6SmokeTest` uses a disposable workbook copy to create a baseline version, change one paragraph, save again, verify the light-terracotta difference marker, delete a paragraph, and restore it. Expected result:

`PASS: version differences and restorable recycle items are working.`

## Requirement traceability

This phase advances AC-060 through AC-067 and AC-151. Full regression coverage for multi-chapter deletion and restoration will be expanded alongside the later relationship and asset workflow tests.
