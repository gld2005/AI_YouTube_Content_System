# Phase 5: Script Presentation and Controls

Phase 5 completes the next operational layer of the English script workspace. New script pages now present the writing model consistently and expose the available paragraph metadata directly in the worksheet.

## Delivered behavior

- A horizontal story-flow strip provides Hook, Context, Tension, Proof, Turning Point, and Resolution prompts.
- The production sidebar has Content, Editing, Management, and AI Preview groups. The sidebar columns can be collapsed and expanded without hiding the document area.
- Paragraph creation supports six types: Narration, On-screen Text, Quote / Dialogue, Visual-only, Music Segment, and Transition.
- Paragraph creation supports seven statuses: To Ideate, Drafting, Finalized, Recorded, Assets Matched, Edited, and Completed.
- New paragraph blocks receive editable Excel list validation for type and status, low-saturation type strips, and status colors.
- `RefreshSelectedParagraphPresentation` reapplies the correct type and status styling after an editable value changes.

## Safety and scope

Paragraph type and status are user-controlled metadata. The workbook does not infer an approval, finalize a paragraph, create assets, or change a project state automatically.

The phase does not yet implement per-group collapse controls, subclaim evidence cards, recycle storage, difference highlighting, or asset synchronization. Those capabilities require their own persistent relationship and approval models.

## Requirement traceability

This phase advances the structural portions of AC-046, AC-050, AC-051, and AC-052. It completes the sidebar-wide collapse behavior while per-group collapse remains open. AC-048 and AC-063 through AC-067 remain open.

## Validation

`RunPhase5SmokeTest` creates a disposable script page, verifies the story flow and four sidebar groups, exercises the sidebar collapse and expansion actions, and applies all six paragraph types. Expected result:

`PASS: story flow, sidebar groups, paragraph types, and statuses are available.`
