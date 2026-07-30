# 03 — Workbook and Data Model Specification

## 1. Logical workbook surfaces

Codex may adjust exact sheet names or split large tables, but the final workbook must provide these logical surfaces:

1. Instructions / 使用说明
2. Dashboard / 总览
3. Content Workspace / 内容工作台
4. Topic Pool / 选题池
5. Video Projects / 视频项目
6. Series Settings / 栏目设置
7. Structure Template Library / 结构模板库
8. Script Index / 脚本总库
9. Script Template / 脚本模板
10. Per-video Current Script pages
11. Per-video Previous Version pages
12. Source Master / 研究来源库
13. Evidence Cards / 证据卡
14. Research Links / 研究关联
15. Packaging Laboratory / 包装实验室
16. Tasks / 任务计划
17. Gantt / 甘特图
18. Project Assets / 项目素材
19. Shared Asset Library / 公共素材库
20. Tag Library / 标签库
21. AI Approval Center / AI审批中心
22. Publish and Optimize / 发布复盘
23. Settings / 设置

## 2. Stable identifiers

Every major record must have a stable ID. IDs must not depend on row number.

Recommended logical IDs:

- Channel ID
- Series ID
- Structure Template ID
- Topic ID
- Video ID
- Chapter ID
- Claim ID
- Paragraph ID
- Source ID
- Evidence ID
- Research Link ID
- Packaging Concept ID
- Task ID
- Asset ID
- Tag ID
- AI Suggestion ID
- Version ID
- Deleted Item ID

## 3. Relationship model

- One channel has many series.
- One series has two to four current structure templates.
- One video belongs to one series or Unclassified.
- One video has many chapters.
- One chapter has one main claim and two to four subclaims.
- One chapter has many paragraphs.
- One paragraph may link to many evidence cards and many assets.
- One source has many evidence cards.
- Sources and evidence can link to many videos, claims, chapters, and paragraphs.
- One video has up to five default packaging concepts, with additional records allowed only if explicitly needed.
- One video has many tasks and assets.
- One project asset may be promoted to the shared library.
- One script has one current version and one previous version.

## 4. Core tables and required fields

### 4.1 Video Projects

Required fields:

- Video ID
- Video Title
- Topic ID
- Series ID
- Primary Project flag
- Content Pillar
- Video Format
- Target Duration
- Target Audience
- Core Promise
- Project-level Status
- Current Working Stage
- Script Status
- Priority
- Owner
- Start Date
- Target Publish Date
- Actual Publish Date
- Total Progress
- Structure Progress
- Writing Progress
- Recording Progress
- Asset Progress
- Editing Progress
- Remaining Days
- Risk Status
- Project Folder Path
- Current Script Sheet
- Previous Version Sheet
- Current Version
- Last Updated
- Notes

### 4.2 Series Settings

Required fields:

- Series ID
- Series Name
- Default flag
- Audience
- Positioning
- Content Promise
- Visual Style
- Target Duration
- Publishing Cadence
- Default Structure Template ID
- Default Folder Template ID
- Folder Trigger Status
- Default Stage Allocation
- Project Transition Thresholds
- Publishing Metric Targets
- Default AI Prompt Set
- Default Task Template ID
- Active flag
- Last Updated

### 4.3 Structure Templates

Required fields:

- Template ID
- Series ID
- Template Name
- Description
- Current Version
- Previous Version reference
- Default flag
- Chapter order and definitions
- Default story flow
- Default paragraph types
- Default transition prompts
- Status
- Last Updated

### 4.4 Script Index

Required fields:

- Video ID
- Video Title
- Series ID
- Current Script Sheet
- Previous Version Sheet
- Current Version
- Script Status
- Last Edited Chapter
- Last Edited Paragraph
- Last Updated
- Total Estimated Duration
- Total Progress
- Pending AI Suggestions
- Recycle Item Count

### 4.5 Chapter records

Required fields:

- Video ID
- Chapter ID
- Sequence
- Chapter Name
- Chapter Objective
- Main Claim
- Viewer State
- Transition to Next Chapter
- Estimated Duration
- Required Fields Complete flag
- Structure Confirmed flag
- Status
- Last Updated

### 4.6 Claims and subclaims

Required fields:

- Video ID
- Chapter ID
- Claim ID
- Claim Type: Main or Subclaim
- Sequence
- Claim Text
- Evidence Summary
- Evidence IDs
- Visual Expression
- Related Paragraph IDs
- Status
- Last Updated

### 4.7 Paragraph records

Required fields:

- Video ID
- Chapter ID
- Paragraph ID
- Sequence
- Paragraph Type
- Paragraph Status
- Estimated Duration
- Actual Timecode
- Current Text
- Previous Text or comparison reference
- Visual Idea
- Evidence IDs or Source IDs
- Asset Requirement
- Subtitle Emphasis
- Music and Sound Notes
- Motion or Fusion Notes
- Change Reason
- Last Updated
- Last AI Action
- Last AI Approval Method
- Deleted flag

### 4.8 Source Master

Required fields:

- Source ID
- Source Type
- Title
- Author or Channel
- Organization or Platform
- Publication Date
- Access Date
- URL or Local Path
- Overall Summary
- Overall Credibility
- Verification Status
- Topic Tags
- Series Tags
- Parse Status
- Last Updated

### 4.9 Evidence Cards

Required fields:

- Evidence ID
- Source ID
- Evidence Classification
- Evidence Text or Summary
- Original Location
- Risk Labels
- Verification Note
- Manual Use Judgment
- Judgment Note
- Supported Claim IDs
- Suggested Paragraph IDs
- Last Updated

### 4.10 Research Links

Required fields:

- Research Link ID
- Source ID
- Evidence ID
- Video ID
- Series ID
- Chapter ID
- Claim ID
- Paragraph ID
- Link Purpose
- Active flag

### 4.11 Packaging Concepts

Required fields:

- Packaging Concept ID
- Video ID
- Variant Number
- Title
- Thumbnail Text
- Thumbnail Visual Description
- Target Viewer
- Click Motivation
- Curiosity or Conflict
- Content Match Score
- Differentiation Score
- Credibility Score
- AI Score
- AI Rationale
- Human Judgment
- Automatic Rank
- Top-Two flag
- Direct Manual Selection flag
- Selected flag
- Impressions
- CTR
- Active Start
- Active End
- Replacement Event
- Before/After Performance
- Final Conclusion
- Last Updated

### 4.12 Tasks

Required fields:

- Task ID
- Video ID
- Source Layer: Base, Series, or Project
- Parent Task ID
- Main Task flag
- Key Child flag
- Stage
- Deliverable
- Description
- Status
- Ready to Confirm flag
- Manual Completion Confirmation
- Priority
- Risk Status
- Weight Mode
- Child Weight
- Suggested Progress
- Confirmed Progress
- Owner
- Planned Start
- Due Date
- Actual Start
- Actual Completion
- Predecessor IDs
- Downstream IDs
- Blocker Reason
- Unblock Condition
- Expected Resolution Date
- Publish Date Impact flag
- Include in Gantt flag
- Last Updated

### 4.13 Project Assets

Required fields:

- Asset ID
- Video ID
- Asset Name
- Asset Type
- File or URL
- Chapter ID
- Paragraph IDs
- Visual Purpose
- Expected Timecode
- Owner
- Status
- Production Notes
- Similar Asset Candidates
- Shared Library Candidate flag
- Last Updated

### 4.14 Shared Asset Library

Required fields:

- Asset ID
- Asset Name
- Asset Type
- File or URL
- Primary Category
- Tag IDs
- Reuse Notes
- Source Project IDs
- Last Used Date
- Active flag

### 4.15 Tags

Required fields:

- Tag ID
- Tag Name
- Tag Category
- Status: Official or Pending Organization
- Synonym Of
- Created Date
- Approved Date
- Active flag

### 4.16 AI Approval records

Required fields:

- AI Suggestion ID
- Video ID
- Chapter ID
- Paragraph ID
- Feature Name
- Original Text
- AI Raw Output
- User-edited Candidate
- Final Adopted Text
- Approval Status
- Adoption Method: Full, Edited, or Excerpt
- Rejection Reason
- Model Provider
- Model Name
- Relevant Parameters
- Created Time
- Reviewed Time
- Reviewer
- Retention Status

## 5. Calculations

### 5.1 Overall script progress

Weighted total:

- Structure: 20%
- Writing Finalization: 35%
- Recording: 15%
- Asset Matching: 15%
- Editing: 15%

### 5.2 Duration weighting

When estimated paragraph durations exist, writing, recording, asset, and editing progress use duration-weighted calculations. Missing duration must not create divide-by-zero errors. Codex must define a documented fallback.

### 5.3 Structure progress

A chapter contributes only when:

- required planning fields pass the completeness check;
- the user manually confirms Structure Confirmed.

### 5.4 Main-task progress

- Calculate suggested progress from child tasks.
- Use equal child weights by default.
- Allow manual weights when selected.
- Do not mark the main task completed until the user confirms.

### 5.5 Risk

Risk combines:

- due-date proximity;
- overdue status;
- blockers;
- dependency impact;
- publish-date impact.

### 5.6 Project-state suggestion

A suggestion is generated only when:

- required conditions are complete;
- progress exceeds the configured threshold;
- there are no unresolved blocking conditions.

## 6. Data validation

Use controlled dropdowns for:

- project status;
- working stage;
- script status;
- paragraph type;
- paragraph status;
- priority;
- task status;
- asset status;
- evidence classification;
- risk labels;
- manual evidence use judgment;
- series;
- structure template;
- AI approval status;
- adoption method;
- tag category and tag status.

## 7. Workbook integrity

- All tables need stable unique names.
- Formulas must avoid hard-coded row limits where feasible.
- Macro actions must preserve IDs and links.
- Sheet names must be sanitized safely.
- Duplicate Video IDs must be blocked or resolved through explicit user choice.
- Broken links, missing sheets, missing local files, and invalid API configuration must produce understandable messages.
- Hidden technical sheets may be used, but user-editable source-of-truth fields must remain accessible.
