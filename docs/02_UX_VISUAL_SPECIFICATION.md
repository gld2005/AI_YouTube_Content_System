# 02 — UX and Visual Specification

## 1. Design intent

The workbook should feel like a warm, focused creator studio with clear operational boundaries. It must not look like an unstyled grid, a crowded accounting sheet, or a cold enterprise dashboard.

## 2. Confirmed visual direction

- Warm gray and beige foundation.
- Low-saturation blue is the primary accent.
- Terracotta orange is the secondary accent.
- Blue should dominate; orange should be used selectively for action, attention, current focus, and change marking.
- Medium information density.
- Card-based fills for summaries and important modules.
- Clear light-gray boundaries.
- Stronger outer borders and lighter inner dividers.
- Rounded dashboard cards combined with stable square-cell data tables.

## 3. Boundary system

### 3.1 Outer boundaries

Major cards, planning blocks, and dashboard sections need a clearly visible warm-gray outline. The outline should be distinct but not dark or heavy.

### 3.2 Inner boundaries

Table cells and internal detail areas use thinner and lighter lines. Internal lines should support scanning without creating a “cage” effect.

### 3.3 Fill behavior

- Dashboard and workspace cards may use full warm off-white, muted blue, or muted terracotta fills.
- Data sheets use a light header fill and restrained alternating row fills.
- Do not fill every small cell with a different color.
- Color must communicate hierarchy or status, not decoration.

## 4. Shape strategy

### Rounded shapes

Use rounded cards for:

- top-level KPI summaries;
- current project card;
- recent projects;
- risk card;
- major dashboard actions.

### Square structures

Use standard square cells and rectangular grouped ranges for:

- project lists;
- source records;
- evidence records;
- tasks;
- asset tables;
- packaging experiments;
- publishing data.

## 5. Dashboard structure

The dashboard is balanced rather than metric-only, flow-only, or action-only.

### Layer 1 — Key cards

Show project title, current status, target publish date, total progress, overdue task count, and asset readiness.

### Layer 2 — Nine-stage progress

Show the nine-stage pipeline with clear stage identity and progress.

### Layer 3 — Action area

Show the most important current actions, due dates, owner, status, and blockers.

## 6. Content workspace structure

### Upper region

A card region containing:

- current primary project;
- two recent projects;
- one at-risk or due-soon project.

The primary project card is visually dominant.

### Lower region

A compact, filterable project list with clear row boundaries and restrained alternating fills.

## 7. Script page structure

### 7.1 Width ratio

- Main document area: approximately 65%.
- Production and management sidebar: approximately 35%.

### 7.2 Top summary

Use a compact card group for the core promise, selected title, Hook, target duration, and current stage.

### 7.3 Detailed planning

Use collapsible grouped sections. Do not display every planning field at full height by default.

### 7.4 Chapter presentation

- Use a light blue chapter title band.
- Preserve visible whitespace between chapters.
- Inside a chapter, keep a continuous document-reading experience.
- Avoid placing every paragraph in an isolated colored card.

### 7.5 Paragraph block

Use a three-row block with a clearly visible but light outer boundary and thin internal separators.

- The paragraph ID sits in a narrow left column using small light-gray text.
- The type label is compact.
- The type color appears mainly as a narrow left strip and small label.
- The narration row has larger, more comfortable body text.
- Production information is visually subordinate and may be collapsed.

## 8. Paragraph type colors

Use six independent low-saturation colors. Exact values may be tuned during implementation, but the categories must remain distinct and calm.

Suggested direction:

- Narration: muted blue
- On-screen Text: terracotta
- Quote / Dialogue: gray-purple
- Visual-only: olive green
- Music Segment: sand-gold
- Transition: warm gray

The main body background remains warm off-white.

## 9. Status colors

Use grouped status families rather than seven unrelated colors:

- To Ideate: warm gray
- Drafting and Finalized: low-saturation blue family
- Recorded, Assets Matched, and Edited: terracotta family
- Completed: low-saturation green

## 10. Typography

- Chinese body text: Microsoft YaHei.
- English, numbers, metrics, and compact labels: Aptos.
- Narration text should be slightly larger and more comfortable than task-table text.
- Paragraph IDs, labels, statuses, and timecodes should be compact.
- Chapter titles should be clearly stronger than body text but not oversized.

Exact sizes can be tuned for desktop Excel, but the hierarchy must remain consistent.

## 11. Density

Use medium density:

- enough whitespace for clarity;
- no oversized empty areas;
- no compressed 15-row dashboard where every cell is equally important;
- body rows should remain practical for daily use.

## 12. Interaction hierarchy

Blue is used for:

- primary information hierarchy;
- active navigation;
- progress;
- selected content;
- writing-stage states.

Terracotta is used for:

- current focus;
- actionable alerts;
- version differences;
- production-stage states;
- approval-required items.

Green is reserved for confirmed completion or successful validation.

## 13. Gridline policy

Native Excel gridlines should be hidden on designed sheets. Required boundaries must be intentionally drawn using the defined border system.

## 14. Accessibility and legibility

- Do not rely on color alone; include labels and text states.
- Avoid low-contrast beige text on beige backgrounds.
- Keep essential values readable at 90% to 100% zoom.
- Use frozen panes where needed.
- Prevent excessive horizontal scrolling on common data sheets.
- The script page may be wider because the 65/35 layout is intentional.
