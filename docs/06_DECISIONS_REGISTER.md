# 06 — Confirmed Decisions Register

This register summarizes decisions explicitly confirmed during product discovery. These are not optional suggestions.

## Visual system

- Warm gray and beige foundation.
- Medium information density.
- Card-based fills.
- Clear boundaries with stronger outer borders and thin internal lines.
- Mixed shapes: rounded dashboard cards, square operational tables.
- Dual accent system: low-saturation blue primary, terracotta orange secondary.
- Blue dominates; orange is selective.
- Microsoft YaHei for Chinese body text; Aptos for English and metrics.

## Dashboard and workspace

- Balanced dashboard: metrics, process, and actions.
- Content Workspace uses a card plus list layout.
- Top cards: one primary, two recent, one at-risk or due-soon, no duplicates.
- Primary project is manual-first with automatic fallback.
- Primary card balances content and progress.

## Script system

- Hybrid script library plus independent per-video script pages.
- Script template is copied for each video.
- New video creation uses VBA.
- New-video input: Video ID and title.
- Missing projects are created automatically.
- Smart defaults are used.
- Current and previous script pages are separate.
- Only one previous version is retained.
- Version save is manual.
- Changed paragraphs are marked light terracotta.
- Paragraph IDs are visible but visually weak.
- Chapter display is a light-blue title band plus document-style body.
- Paragraphs use a three-row small block.
- Paragraph status spans writing through production.
- Paragraph status colors use grouped families.
- Top planning uses a summary card plus collapsible full planning.
- Story outline uses a horizontal flow plus detailed chapter list.
- Script layout is 65% document and 35% production sidebar.
- Chapters and paragraphs are dynamically added by macro.
- New paragraphs are inserted after the selected paragraph.
- Deleted items go to a recycle area.
- Recycle area retains 20 items.
- Six paragraph types are supported.
- Each paragraph type has an independent low-saturation color label and strip.
- Paragraph typography is mixed: reading-oriented body plus compact metadata.
- Sidebar groups are Content, Editing, Management, and AI preview.
- AI output always requires user approval before merge.
- AI merge supports full, edited, and excerpt adoption.
- AI change logging is compact on script pages and complete in the approval center.

## Claims and research

- Full planning area is required.
- Chapter planning uses main claim, evidence, and visual expression.
- Each chapter has one main claim and two to four subclaims.
- Subclaims use compact argument cards.
- Argument cards are collapsed by default.
- Research uses Source Master plus Evidence Cards plus Research Links.
- Sources support dual attribution across videos and reusable topics.
- Evidence uses detailed classification and risk labels.
- Evidence usage remains a manual judgment.
- Research import is hybrid: extraction plus AI plus approval plus manual fallback.
- All major source types are in scope with tiered support.

## Packaging

- Script page shows the selected packaging only.
- Packaging Laboratory stores all alternatives.
- Five default concepts per video.
- Full planning-to-performance closed loop.
- Two-stage automatic screening plus final human selection.
- Direct manual selection remains available and overrides rankings.

## AI and local assistant

- AI is optional at the workbook level but directly callable when configured.
- Provider presets plus full custom configuration.
- Environment-variable key first, session input fallback.
- Modular AI feature enablement.
- Default features match the specification.
- All AI results require approval.
- Script preview plus centralized approval center.
- Local-first architecture with optional cloud replacement.
- Windows executable plus Python source.
- Excel checks service status and offers one-click start.
- Optional Windows autostart.

## Channel and series

- One channel with multiple series.
- New videos use default/recent series and can be Unclassified.
- Inheritance choices are selectable.
- Default inheritance is standardized except packaging.
- Each series has two to four structure templates.
- Template selection combines default, AI recommendation, and human choice.
- AI defaults to recommending and modifying existing templates.
- Free structure generation requires explicit action.
- Accepted new structures prompt Save Current Only, Save New Template, or Overwrite.
- Structure templates keep current and previous versions only.

## Projects, tasks, and schedule

- Program and project storage are separated.
- Project roots can be on any disk and are relocatable.
- Project folder templates are configurable by series.
- Folder creation is triggered by status and confirmed by the user.
- Default trigger is Planning.
- Project status uses a two-layer model.
- Status updates are suggestions plus human confirmation.
- Suggestions use conditions, thresholds, and blocker checks.
- Thresholds are series-configurable with defaults.
- Tasks use Base, Series, and Project layers.
- Tasks are hierarchical.
- Main-task progress is computed but completion is manually confirmed.
- Child weights are equal by default and manually adjustable when needed.
- Priority uses P0/P1/P2 plus automatic deadline and blocker risk.
- Dependencies use the complete model.
- Timeline, task list, and Gantt are all required.
- Forward and reverse planning are both supported.
- Reverse planning uses series defaults, AI suggestion, and human approval.
- Schedule granularity is main tasks plus key child tasks.
- Scheduling uses natural days.
- No effort-hour tracking.

## Assets and tags

- Project Assets plus Shared Asset Library.
- Project assets use the production-oriented field set.
- Asset status is only To Find, Found, Used.
- Asset source method is not required.
- Finalized paragraphs prompt asset synchronization.
- Duplicate detection combines text matching, AI semantic comparison, and human choice.
- Published projects trigger reusable-asset recommendations.
- Shared assets use primary type plus categorized tags.
- Tags use fixed categories with custom additions pending organization.
