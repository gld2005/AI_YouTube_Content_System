# 01 — Product Specification

## 1. Product intent

The product is a creator-focused Excel workspace for planning, researching, scripting, producing, publishing, and improving YouTube videos. It must support one channel with multiple recurring series or content pillars.

The system is not a generic project tracker. The script, evidence, claims, packaging hypotheses, and production handoff are the primary objects. Dates, tasks, dependencies, and progress exist to make content execution reliable.

## 2. Lifecycle

The canonical lifecycle is:

1. Discover
2. Research
3. Script
4. Assets
5. Edit
6. Motion
7. Post
8. Publish
9. Optimize

The workbook must preserve this nine-stage mental model while also supporting a simpler project-level status model.

## 3. Project status model

### 3.1 Project-level status

- Topic Selection
- Planning
- Production
- Publishing
- Completed

The user-facing Chinese labels should be:

- 选题中
- 策划中
- 制作中
- 发布中
- 已完成

### 3.2 Current working stage

The current stage is tracked separately:

- Research
- Script
- Assets
- Edit
- Motion
- Post
- Publish
- Optimize

### 3.3 Transition behavior

- The system may recommend a project-level status transition.
- The recommendation requires critical conditions to be complete, a configurable progress threshold to be met, and no unresolved blocking item.
- The user must confirm every transition.
- Default thresholds are 60%, 85%, 90%, and 100% for the four transitions.
- Each series may override the thresholds.

## 4. Channel and series model

The workbook manages one channel with multiple series.

Each series can define:

- series name and identifier;
- audience and positioning;
- content promise;
- visual language;
- target duration;
- publishing cadence;
- target performance metrics;
- default folder template;
- folder-creation trigger status;
- project-status thresholds;
- default task template;
- default AI prompt set;
- two to four script structure templates.

When creating a video:

- the default or most recently used series is preselected;
- the user may change it or choose “Unclassified”;
- inheritance options are shown;
- default selected inheritance items are series positioning, chapter structure, visual specification, AI prompts, production tasks, and publishing targets;
- packaging templates are not selected by default.

## 5. Structure templates

Each series supports two to four structure templates.

- One template can be the default.
- AI may recommend a template based on topic, format, target duration, and core thesis.
- The user makes the final choice.
- By default, AI may recommend and adjust an existing template.
- A separate “Free Structure Design” action is required before AI may propose a completely new structure.
- When an AI-generated structure is approved, the user chooses whether it is current-video-only, saved as a new series template, or used to overwrite an existing template.
- Each structure template keeps only the current version and one previous version.

## 6. Video creation

A macro-enabled “New Video Script” action must request:

- Video ID;
- video title.

It must also allow the user to confirm or change the series and inheritance options.

If the Video ID does not exist, the system creates the project record automatically.

Default values for a newly created project:

- project status: Planning / 策划中;
- script status: Not Started;
- priority: P1;
- start date: current date;
- last updated: current date;
- current script version: V1.

Audience, target duration, and publish date remain blank unless inherited.

The system must create:

- a project record;
- a script index record;
- a current script sheet copied from the script template;
- a paired previous-version sheet or a prepared equivalent that becomes active after the first version save.

## 7. Content workspace

The content workspace is the daily operating surface.

### 7.1 Upper card area

Show four non-duplicated project cards:

- one current primary project;
- two recently edited projects;
- one upcoming-risk or due-soon project.

Primary project selection:

- use the manually selected primary project when available;
- otherwise select the highest-priority active project automatically.

The primary card balances content and execution information:

- core content promise;
- current title;
- current stage;
- script completion;
- most recently edited chapter;
- next action;
- deadline and risk indicator.

### 7.2 Lower list area

Provide a compact list of all projects with filtering, status, series, script links, update time, risk, and publish date.

## 8. Script workspace model

### 8.1 Per-video pages

The workbook contains a reusable script template. Every approved video receives:

- a dedicated current script page;
- a dedicated previous-version page.

The current and previous pages use naming similar to:

- `YT-001｜脚本`
- `YT-001｜上一版`

Sheet-name sanitization and uniqueness must be handled safely.

### 8.2 Page layout

The script page uses:

- a document-style main writing area occupying approximately 65% of the usable width;
- a production support area occupying approximately 35%;
- collapsible groups in the support area.

### 8.3 Top planning area

A fixed summary card displays:

- core promise;
- current selected title;
- Hook;
- target duration;
- current stage.

A collapsible detailed planning area includes:

- topic and thesis;
- target audience;
- click motivation;
- viewer takeaway;
- title direction;
- Hook direction;
- chapter outline;
- opening emotion;
- central tension or suspense;
- information and emotional rhythm;
- turning point;
- ending aftertaste;
- competitor differentiation;
- visual tone;
- reference works;
- required content;
- explicitly excluded content;
- post-publication validation metrics.

### 8.4 Story flow and chapter outline

Use a two-level presentation:

- a horizontal high-level story flow;
- a collapsible detailed chapter list.

Each chapter contains:

- chapter objective;
- main claim;
- two to four subclaims;
- evidence references;
- visual expression;
- viewer cognitive or emotional state;
- transition sentence to the next chapter;
- estimated duration;
- structure-complete check;
- manual structure confirmation.

A chapter counts toward structure completion only after required fields are complete and the user confirms it.

### 8.5 Claims and argument cards

Each chapter has one main claim and two to four subclaims.

Each subclaim uses a compact argument card containing:

- subclaim;
- evidence summary;
- Source ID or Evidence ID references;
- visual expression;
- related script paragraphs;
- status.

Argument cards are collapsed by default to show subclaim, source reference, and status. Expanded view shows evidence, visuals, and paragraph links.

### 8.6 Paragraph hierarchy

The hierarchy is:

**Video → Chapter → Paragraph → optional Shot detail**

A paragraph is the normal writing unit. Shot-level detail is added only where needed.

Every paragraph has a stable, lightly displayed identifier such as `01-03`.

Each paragraph occupies a small three-row block:

1. identifier, type, status, and estimated duration;
2. narration or primary content;
3. visual and production information.

Production information can be collapsed.

### 8.7 Paragraph types

Supported types:

- Narration
- On-screen Text
- Quote / Dialogue
- Visual-only
- Music Segment
- Transition

Each type has its own low-saturation label and left-side color strip. The main paragraph background remains warm off-white.

### 8.8 Paragraph status

Supported statuses:

- To Ideate
- Drafting
- Finalized
- Recorded
- Assets Matched
- Edited
- Completed

Color behavior:

- early/unstarted: warm gray;
- writing and finalized: low-saturation blue family;
- production statuses: terracotta family;
- completed: low-saturation green.

### 8.9 Right-side production groups

The support area contains collapsible groups:

- Content: visual idea, Source IDs, asset requirements;
- Editing: subtitle emphasis, music and sound, animation or Fusion notes, timecodes;
- Management: status, change reason, last update;
- AI preview: current suggestion, comparison, and approval controls.

The Content group is open by default. Other groups may be collapsed until needed.

## 9. Script versioning and deletion

### 9.1 Version rule

Only two script versions exist per video:

- current version;
- previous version.

A “Save New Version” action:

- copies the complete current script into the previous-version page;
- overwrites the older previous version;
- preserves stable paragraph IDs;
- records version date and reason when provided.

Opening or closing the workbook must not create a version automatically.

### 9.2 Difference marking

Paragraphs that differ from the previous version are marked with a light terracotta highlight. Unchanged paragraphs remain neutral.

### 9.3 Recycle area

Deleted chapters and paragraphs move to a collapsed recycle area instead of being destroyed immediately.

- Keep the most recent 20 deleted items per script.
- Store original chapter, paragraph ID, deletion time, content, and metadata.
- Allow restoration.
- When the limit is exceeded, remove the oldest item.
- Saving a new version does not clear the recycle area.

## 10. Script progress

The primary card shows total progress plus five sub-progress metrics:

- Structure
- Finalized Writing
- Recording
- Asset Matching
- Editing

Total progress weights:

- Structure: 20%
- Finalized Writing: 35%
- Recording: 15%
- Asset Matching: 15%
- Editing: 15%

Paragraph-based progress should use estimated paragraph duration rather than equal paragraph counts when duration exists. The system must handle missing duration safely.

## 11. Research system

### 11.1 Architecture

Use three logical layers:

- Source Master
- Evidence Cards
- Research Links

A source is stored once and can be reused across videos.

### 11.2 Source Master

Store:

- Source ID;
- title;
- author or channel;
- organization or platform;
- publication date;
- access date;
- URL or local path;
- source type;
- overall summary;
- overall credibility;
- verification status;
- topic and series tags.

### 11.3 Evidence Cards

A source can produce multiple evidence cards. Each card stores:

- Evidence ID;
- Source ID;
- fact, statistic, quote, opinion, or inference;
- original text excerpt or precise summary;
- page, timestamp, or location;
- evidence classification;
- risk labels;
- verification note;
- supported claim or subclaim;
- possible paragraph use;
- manual use judgment;
- judgment note.

Evidence classifications:

- Verifiable Fact
- Statistical Data
- First-person Quote
- Expert Opinion
- Author Judgment
- Creator Inference

Risk labels:

- Cross-validated
- Single Source
- Disputed
- Possibly Outdated
- Needs More Evidence
- Not Suitable for Direct Use

The system does not block usage automatically. Final evidence usage is a manual judgment.

Manual use judgment options:

- Safe to Use
- Use with Careful Wording
- Do Not Use Yet

### 11.4 Research Links

Research links connect sources and evidence to:

- one or more Video IDs;
- topics;
- series;
- chapters;
- main claims;
- subclaims;
- paragraph IDs.

## 12. Research import

Research import is hybrid:

- extract metadata automatically when possible;
- use AI to propose summaries and evidence cards;
- require approval before writing AI-generated research content;
- allow manual entry when parsing fails.

Source coverage is tiered:

### Full automation priority

- standard web pages;
- news articles;
- YouTube;
- PDFs;
- local documents.

### Semi-automatic support

- Bilibili;
- Reddit;
- X / Twitter;
- podcasts and audio sources;
- platforms requiring login or blocked extraction.

When direct parsing fails, the user can paste text, subtitles, or an exported file.

## 13. Packaging laboratory

### 13.1 Relationship to script

The script summary shows only the currently selected title and thumbnail direction. All alternatives live in the Packaging Laboratory.

### 13.2 Default variants

Each video supports five default title and thumbnail concepts.

Each concept stores planning-stage information:

- title;
- thumbnail text;
- thumbnail visual description;
- target viewer;
- click motivation;
- curiosity or conflict mechanism;
- content match;
- differentiation;
- AI score and rationale;
- human judgment;
- selected status.

After publication, the same record expands to include:

- impressions;
- CTR;
- active time range;
- replacement history;
- before-and-after performance;
- final conclusion.

### 13.3 Selection behavior

- Scoring and AI recommendation identify the top two concepts.
- The user chooses the final concept.
- A direct manual selection option always exists and overrides automatic ranking.

## 14. Task system

### 14.1 Task sources

Tasks come from three layers:

- universal base tasks;
- series-specific tasks;
- project-specific temporary tasks.

### 14.2 Hierarchy

- Main tasks represent deliverables.
- Child tasks represent operational steps.
- Main tasks appear by default.
- Child tasks are collapsible.

### 14.3 Completion

- Child completion generates a suggested main-task progress percentage.
- Child tasks are equal-weight by default.
- Complex tasks may switch to manually assigned child weights.
- Even when all child tasks are complete, the main task becomes “Ready to Confirm,” not automatically “Completed.”
- The user confirms final completion.

### 14.4 Priority and risk

Priority:

- P0
- P1
- P2

Automatic risk indicators:

- Normal
- Due Soon
- Overdue
- Blocked

### 14.5 Dependencies

Track:

- predecessor tasks;
- downstream affected tasks;
- blocker reason;
- unblock condition;
- expected resolution date;
- whether the publication date is affected.

## 15. Timeline and scheduling

Use a combined presentation:

- a simple stage timeline on the content workspace;
- a detailed task list;
- a dedicated Gantt view for main tasks and selected key child tasks.

Scheduling supports:

- forward planning from a start date;
- reverse planning from a publish date.

If a publish date exists, reverse planning is the default. Otherwise, forward planning is the default.

Reverse planning behavior:

- load the series default stage allocation;
- allow AI to suggest adjustments based on duration, research complexity, asset difficulty, and motion intensity;
- require user approval;
- allow manual date editing.

Planning uses natural calendar days. It does not model work hours or effort estimates.

## 16. Asset system

### 16.1 Two libraries

- Project Assets
- Shared Asset Library

### 16.2 Project asset fields

- Asset ID;
- asset name;
- asset type;
- file or URL;
- Video ID;
- chapter and paragraph links;
- visual purpose;
- expected use timecode;
- owner;
- current status;
- production notes.

Asset status is intentionally simple:

- To Find
- Found
- Used

The system does not require the user to classify whether an asset is externally sourced, AI-generated, or self-created.

### 16.3 Script-to-asset synchronization

When a paragraph becomes Finalized, the system summarizes its asset requirements and asks whether to add them to Project Assets.

When a similar asset already exists:

- use name and keyword matching first;
- optionally use AI semantic comparison for likely matches;
- let the user choose Link Existing, Create New, or Merge Description.

### 16.4 Shared asset promotion

After publication, the system recommends reusable Used assets. The user confirms which items enter the Shared Asset Library.

### 16.5 Shared asset classification

Primary asset categories:

- Video
- Image
- Music
- Sound Effect
- Icon
- Chart
- Animation Template

Tags use fixed categories with controlled custom additions. New tags enter a “Pending Organization” state before becoming official.

Fixed tag categories include:

- Person
- Team / Organization
- Topic
- Emotion
- Use Case
- Visual Style
- Scene

## 17. Publishing and optimization

The Publish and Optimize area must record:

- platform;
- publish date;
- final title;
- thumbnail version;
- duration;
- impressions;
- views;
- CTR;
- watch time;
- average view duration;
- average percentage viewed;
- likes;
- comments;
- subscribers gained;
- 24-hour, 7-day, and 28-day views;
- first-30-second retention;
- largest drop point;
- largest replay point;
- comment themes;
- strengths;
- problems;
- next experiment.

It should close the loop with the Packaging Laboratory, structure templates, reusable assets, and future series decisions.

## 18. Navigation

Use a combined navigation model:

- Dashboard: major feature entry points and portfolio health;
- Content Workspace: all projects and script entry points;
- Script pages: only essential actions such as Return to Workspace, Save New Version, Add Chapter, and Add Paragraph.

Avoid repeating a large navigation bar on every sheet.
