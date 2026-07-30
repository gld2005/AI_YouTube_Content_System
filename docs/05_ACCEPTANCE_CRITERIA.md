# 05 — Acceptance Criteria

## A. Scope and governance

**AC-001** Codex provides a plan and receives explicit approval before implementation begins.

**AC-002** Every requirement group in the traceability document is mapped to an implemented feature or an approved equivalent.

**AC-003** The reference workbook remains preserved and is not overwritten.

**AC-004** All technical documentation, code comments, logs, configuration keys, and tests are in English.

**AC-005** User-facing workbook labels and messages are in Simplified Chinese unless changed by the user.

## B. Visual experience

**AC-010** The workbook uses a warm gray and beige base, low-saturation blue as the primary accent, and terracotta as the secondary accent.

**AC-011** Major modules have visibly stronger warm-gray outer boundaries and lighter internal separators.

**AC-012** Native gridlines are hidden on designed sheets.

**AC-013** Dashboard and workspace summaries use rounded cards; operational tables remain stable square-cell structures.

**AC-014** The workbook maintains medium information density and does not resemble an unformatted full-grid sheet.

**AC-015** Chinese body text uses Microsoft YaHei where available, and English/numeric content uses Aptos where available.

**AC-016** Essential information remains legible at normal desktop zoom.

## C. Dashboard and content workspace

**AC-020** The Dashboard contains KPI cards, nine-stage progress, and a current-action area.

**AC-021** The Content Workspace shows one primary, two recent, and one at-risk/due-soon project without duplicates.

**AC-022** A manually selected primary project overrides automatic selection.

**AC-023** Without a manual primary project, the highest-priority active project is selected automatically.

**AC-024** The primary card shows content promise, current title, current stage, script progress, recent chapter, next action, deadline, and risk.

**AC-025** The lower workspace list links to project and script records and supports practical filtering.

## D. Series and templates

**AC-030** The system supports one channel with multiple series.

**AC-031** Each series supports two to four script structure templates.

**AC-032** New videos preselect a default or recent series, allow replacement, and allow Unclassified.

**AC-033** New-video inheritance defaults match the specification and packaging inheritance is off by default.

**AC-034** AI can recommend and adjust existing templates only by default.

**AC-035** A separate explicit action is required for AI to create a completely new structure.

**AC-036** After approving a new structure, the user can choose current-video-only, save as new, or overwrite existing.

**AC-037** Each structure template retains only a current and previous version.

## E. Video creation and scripts

**AC-040** “New Video Script” requests Video ID and title and allows series/inheritance confirmation.

**AC-041** A missing project record is created automatically with the confirmed smart defaults.

**AC-042** The action creates a current script page, previous-version page or prepared equivalent, script index, project record, and working links.

**AC-043** Duplicate Video IDs are not created silently.

**AC-044** Each script page uses approximately 65% document area and 35% production sidebar.

**AC-045** The top summary card and detailed collapsible planning fields match the product specification.

**AC-046** A horizontal story flow and detailed collapsible chapter list are both present.

**AC-047** Chapters support one main claim and two to four subclaims.

**AC-048** Subclaim cards collapse to summary and expand to evidence, visuals, and paragraph links.

**AC-049** Paragraphs use stable IDs, a three-row block, a compact type label, and a left-side type strip.

**AC-050** All six paragraph types are available and visually distinguishable without changing the main warm off-white body background.

**AC-051** All seven paragraph statuses are available with grouped status colors.

**AC-052** Content, Editing, Management, and AI preview groups are available in the right sidebar and can be collapsed.

## F. Script versioning and recycle area

**AC-060** Each video retains only current and previous script versions.

**AC-061** Saving a new version is an explicit action and never occurs automatically on open or close.

**AC-062** The previous page is overwritten with the current script when the user saves a new version.

**AC-063** Changed paragraphs are detected by stable ID and marked in light terracotta.

**AC-064** Unchanged paragraphs remain neutral.

**AC-065** Deleted chapters and paragraphs enter a restorable recycle area.

**AC-066** The recycle area retains the most recent 20 deleted items and removes the oldest when exceeded.

**AC-067** Saving a new version does not clear the recycle area.

## G. Progress and status

**AC-070** The five progress metrics and 20/35/15/15/15 total weighting are implemented.

**AC-071** Paragraph production progress is duration-weighted when duration exists and handles missing duration safely.

**AC-072** Structure progress requires both field completeness and manual confirmation.

**AC-073** Project-level status and current working stage are separate.

**AC-074** Project-level transitions are suggestions until user confirmation.

**AC-075** Suggestions require critical conditions, threshold completion, and no blocker.

**AC-076** Default thresholds are 60%, 85%, 90%, and 100%, with series overrides.

## H. Research

**AC-080** Sources are stored once in Source Master and can be reused across videos.

**AC-081** A source can contain multiple evidence cards.

**AC-082** Research links connect sources or evidence to videos, series, chapters, claims, and paragraphs.

**AC-083** All six evidence classifications and six risk labels are available.

**AC-084** Evidence usage remains a manual decision with the three specified judgment values.

**AC-085** Standard web, news, YouTube, PDF, and local documents are prioritized for automatic parsing.

**AC-086** Bilibili, Reddit, X, podcasts, audio, and restricted sources support best-effort parsing with manual fallback.

**AC-087** AI-generated summaries and evidence cards require approval before becoming authoritative.

## I. Packaging

**AC-090** Each video starts with five packaging concepts.

**AC-091** The script summary displays only the selected title and thumbnail direction.

**AC-092** The Packaging Laboratory stores planning hypotheses and post-publication metrics.

**AC-093** Scoring and AI recommendation identify a top two.

**AC-094** The user chooses the final concept.

**AC-095** Direct manual selection can override automatic ranking.

## J. Tasks, dependencies, and schedule

**AC-100** Tasks support Base, Series, and Project source layers.

**AC-101** Main tasks and collapsible child tasks are supported.

**AC-102** Child tasks calculate suggested progress; main tasks require manual completion confirmation.

**AC-103** Equal weights are the default and manual weighting is available.

**AC-104** P0, P1, and P2 priorities and Normal, Due Soon, Overdue, and Blocked risks are available.

**AC-105** Full dependency fields are available, including downstream impact and publish-date impact.

**AC-106** The system provides a workspace timeline, detailed task list, and Gantt view.

**AC-107** Forward and reverse planning are supported.

**AC-108** Reverse planning uses series defaults, optional AI suggestions, user approval, and manual adjustment.

**AC-109** Main tasks and selected key child tasks appear in the Gantt by default.

**AC-110** Scheduling uses natural days and does not require effort-hour entry.

## K. Assets and tags

**AC-120** Project Assets and a Shared Asset Library both exist.

**AC-121** Project Assets use the confirmed production-oriented fields and simple To Find, Found, Used statuses.

**AC-122** Finalizing a paragraph triggers an optional asset-sync prompt.

**AC-123** Similar-asset detection uses text matching and optional AI semantic comparison.

**AC-124** The user chooses Link Existing, Create New, or Merge Description.

**AC-125** After publication, the system recommends reusable Used assets for manual promotion.

**AC-126** Shared assets use fixed primary categories and categorized tags.

**AC-127** Custom tags enter Pending Organization before becoming official.

## L. AI approval and security

**AC-130** Provider presets and full custom OpenAI-compatible configuration are available.

**AC-131** API keys are read from environment variables first and may be entered for the session as a fallback.

**AC-132** Plaintext API keys are not stored in the workbook, source repository, or logs.

**AC-133** Default AI actions match the specified feature set and can be enabled or disabled.

**AC-134** Every AI result enters approval before modifying source-of-truth content.

**AC-135** The script sidebar and central AI Approval Center both display pending suggestions.

**AC-136** Full Adoption, Edit Before Adoption, and Excerpt Adoption are supported.

**AC-137** Analysis actions never modify script text directly.

**AC-138** Compact audit information appears on the script page and full audit data appears in the AI Approval Center.

**AC-139** Adopted and pending AI records are retained as specified; rejected items keep at least the rejection reason.

## M. Local assistant and folders

**AC-140** A Windows executable and Python source mode are both delivered.

**AC-141** Excel checks assistant availability and provides a one-click start option.

**AC-142** Windows autostart is optional and not forced.

**AC-143** Program and project storage are separated.

**AC-144** Project roots can reside on any disk and can be reselected after relocation.

**AC-145** Series-specific project folder templates are supported.

**AC-146** Folder creation is triggered by configurable project status, defaults to Planning, and requires confirmation.

**AC-147** The architecture can switch to a cloud-compatible service without changing the core workbook behavior.

## N. Reliability

**AC-150** AI, parsing, macro, path, duplicate ID, and missing-sheet failures provide actionable messages.

**AC-151** No failure silently overwrites approved script, research, packaging, or project content.

**AC-152** Key formulas have no obvious `#REF!`, `#DIV/0!`, `#VALUE!`, or broken-name errors in normal use.

**AC-153** Core macros are testable on a working copy and preserve stable IDs and relationships.

**AC-154** The workbook opens without repair warnings in the target desktop Excel environment.
