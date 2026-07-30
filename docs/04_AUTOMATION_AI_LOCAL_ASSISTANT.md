# 04 — Automation, AI, and Local Assistant Specification

## 1. Architecture principle

The solution is local-first with optional cloud compatibility.

- Excel is the primary interface, planning environment, and approval surface.
- VBA handles workbook-native interactions, sheet creation, navigation, versioning, grouping, validation, and orchestration.
- A local assistant handles parsing, model calls, richer text processing, file inspection, and platform-specific extraction.
- The local assistant is delivered both as a Windows executable and as Python source.
- A future cloud service may replace the local endpoint if it implements the same functional contract.

Codex may choose the exact protocol and internal module boundaries, but the user-visible behavior is mandatory.

## 2. VBA capabilities

The workbook must provide reliable actions for:

### 2.1 New video creation

- Request Video ID and video title.
- Show or confirm the series.
- Show inheritance options.
- Validate duplicate IDs.
- Create a project record if absent.
- Apply smart defaults.
- Create the current script sheet.
- Create or prepare the previous-version sheet.
- Add the script index record.
- Create links and navigation.
- Evaluate whether project-folder creation should be suggested.

### 2.2 Script structure actions

- Add Chapter.
- Delete Chapter to recycle area.
- Restore Chapter.
- Add Paragraph after the currently selected paragraph.
- Delete Paragraph to recycle area.
- Restore Paragraph.
- Recalculate sequence and stable display numbering without changing stable IDs.
- Maintain chapter, claim, evidence, asset, and task links.

### 2.3 Version action

“Save New Version” must:

- copy the complete current script to the previous-version page;
- overwrite the former previous version;
- record the action time;
- compare paragraphs by stable ID;
- mark changed paragraphs in the current script with light terracotta;
- leave unchanged paragraphs neutral;
- never run automatically on workbook open or close.

### 2.4 Navigation

Provide:

- major navigation actions on the Dashboard;
- script-entry links from the Content Workspace;
- Return to Workspace, Save New Version, Add Chapter, and Add Paragraph actions on script pages.

### 2.5 Asset synchronization

When a paragraph becomes Finalized:

- collect its asset requirements;
- ask whether to create Project Asset records;
- detect possible duplicates using fast text matching;
- optionally request AI semantic comparison;
- show Link Existing, Create New, and Merge Description choices.

### 2.6 Task and status automation

- Recalculate suggested task progress.
- Flag Ready to Confirm.
- Recalculate risk.
- Generate project-state transition suggestions.
- Never confirm project transitions automatically.
- Maintain the stage timeline and Gantt data source.

### 2.7 Project folder creation

- Use the series folder template.
- Trigger a suggestion when the configured project status is reached.
- Require user confirmation.
- Support project roots on any disk.
- Support relocation and path re-selection.
- Avoid unnecessary C-drive use.

## 3. AI provider configuration

Support provider presets and full custom configuration.

Preset targets:

- OpenAI
- DeepSeek
- OpenRouter
- Alibaba Cloud Model Studio / compatible endpoint
- Custom OpenAI-compatible service

Configurable fields:

- API base URL
- model name
- environment variable name
- request format or compatibility mode
- temperature
- maximum output length
- timeout
- optional custom headers where safe

## 4. API key security

- Read from a Windows environment variable first.
- If unavailable, request temporary session input.
- Do not save plaintext keys in cells, hidden sheets, VBA modules, configuration files committed to the project, or logs.
- Do not expose keys in error messages.

## 5. Default AI feature set

The settings surface allows individual AI actions to be enabled or disabled.

Default enabled features:

### Script core

- Generate Chapter Outline
- Generate Hook
- Expand Current Paragraph
- Polish Current Paragraph
- Compress Current Paragraph
- Rewrite Current Paragraph

### Quality analysis

- Check Logic Gaps
- Check Repetition and Filler
- Check Missing Sources
- Analyze Rhythm and Retention Risk
- Generate Overall Revision Suggestions

### Production handoff

- Generate Visual Ideas
- Generate Asset List
- Estimate Paragraph Duration

### Research and packaging

- Generate Research Questions
- Generate Title and Thumbnail Concepts

Other supported actions may be enabled through Settings.

## 6. AI approval policy

Every AI output requires explicit user approval before it can change source-of-truth content.

### 6.1 Display model

- The active script shows a right-side preview for the current suggestion.
- All unprocessed suggestions also enter the AI Approval Center.
- The user may approve immediately or review later.

### 6.2 Adoption methods

Support:

- Full Adoption
- Edit Before Adoption
- Excerpt Adoption

### 6.3 Behavior by feature type

- Generation features: preview and choose replace, append, write to notes, or reject.
- Rewrite features: show original and AI candidate side by side.
- Analysis features: write suggestions to the analysis area; never change the script directly.

### 6.4 Audit model

The script page shows a compact audit line:

- paragraph ID;
- AI feature;
- adoption method;
- processed time.

The AI Approval Center stores the full record:

- original text;
- raw model output;
- user-edited candidate;
- final adopted text;
- adoption method;
- model and parameters;
- reason and timestamps.

Retention:

- pending suggestions: keep full content;
- adopted suggestions: keep full content and final result;
- rejected suggestions: keep the rejection reason; full rejected content may be cleaned periodically.

## 7. Research parsing

### 7.1 Full automation priority

Attempt direct extraction for:

- web pages;
- news articles;
- YouTube pages and subtitles;
- PDFs;
- local documents.

### 7.2 Semi-automatic support

Attempt extraction where possible for:

- Bilibili;
- Reddit;
- X / Twitter;
- podcasts;
- audio sources;
- login-restricted sources.

If extraction fails, request pasted text, subtitles, or an exported file.

### 7.3 Approval behavior

Parsed metadata may be written after validation. AI-generated summaries, evidence cards, risk labels, or claim links must enter approval before becoming authoritative records.

## 8. Local assistant distribution

Provide:

- a Windows executable for ordinary use;
- Python source for debugging and extension.

The local assistant must support:

- a health or status check;
- startup from a user action in Excel;
- configurable path discovery;
- structured requests and responses;
- safe timeout and error behavior;
- local file access only when explicitly requested;
- provider configuration without storing plaintext secrets.

Excel opening behavior:

- check whether the local assistant is available;
- if unavailable, show a one-click start action;
- allow optional user-configured Windows autostart;
- do not force background startup.

## 9. Program and project storage

Program and project data are separated.

- The assistant has a stable tool directory.
- The workbook may reside on any drive.
- Project files may use a user-selected project root on any drive.
- The workbook remembers the active project root.
- Relocation must be recoverable through path re-selection.

## 10. Project folder templates

Folder templates are configurable by series.

A default full-production template may contain logical areas equivalent to:

- Planning
- Research
- Script
- Assets
- Edit
- Exports
- Publish

Series may define different structures.

Creation is triggered by project status, defaults to Planning, and always requires user confirmation.

## 11. Scheduling and AI

For reverse planning:

- load the series allocation;
- allow AI to recommend adjustments;
- explain the reason for adjustments;
- require approval;
- update main tasks and selected key child tasks only;
- use natural calendar days;
- do not model effort hours.

## 12. Error handling

Provide actionable, non-technical user messages in Chinese for workbook users while keeping technical logs in English.

Handle at minimum:

- local assistant unavailable;
- API key unavailable;
- provider or model error;
- timeout;
- unsupported or blocked source;
- missing subtitle;
- invalid local file;
- duplicate Video ID;
- invalid or duplicate sheet name;
- missing script or previous-version page;
- broken project root;
- permission failure when creating folders;
- partial AI response;
- malformed structured output.

No failure may silently overwrite approved content.
