# AI YouTube Content System — Codex Starter Package

This package is the source-of-truth handoff for building a macro-enabled Excel content planning and production system supported by a local AI and research assistant.

## What this system is
It manages the complete content lifecycle:

**Discover → Research → Script → Assets → Edit → Motion → Post → Publish → Optimize**

The core value is content planning, research, argument construction, long-form script writing, and the transition from script to production. Project tracking exists to support the content workflow, not replace it.

## Package contents

- `AGENTS.md` — non-negotiable operating rules for Codex.
- `CODEX_INITIAL_PROMPT.md` — a ready-to-use initial instruction.
- `docs/01_PRODUCT_SPECIFICATION.md` — complete product behavior and workflows.
- `docs/02_UX_VISUAL_SPECIFICATION.md` — visual language and workbook experience.
- `docs/03_WORKBOOK_DATA_MODEL.md` — entities, sheets, IDs, relationships, and calculations.
- `docs/04_AUTOMATION_AI_LOCAL_ASSISTANT.md` — VBA, AI, research parsing, local service, and security requirements.
- `docs/05_ACCEPTANCE_CRITERIA.md` — testable definition of done.
- `docs/06_DECISIONS_REGISTER.md` — concise register of all confirmed decisions.
- `docs/07_REQUIREMENTS_TRACEABILITY.md` — requirement-to-acceptance mapping.
- `reference/AI_YouTube_Content_Pipeline_Baseline.xlsx` — the current workbook prototype for reference and migration testing.
- `reference/BASELINE_NOTES.md` — how to interpret the prototype.

## Source-of-truth precedence
If anything appears inconsistent, use this order:

1. `AGENTS.md`
2. `docs/06_DECISIONS_REGISTER.md`
3. `docs/01_PRODUCT_SPECIFICATION.md`
4. `docs/02_UX_VISUAL_SPECIFICATION.md`
5. `docs/03_WORKBOOK_DATA_MODEL.md`
6. `docs/04_AUTOMATION_AI_LOCAL_ASSISTANT.md`
7. `docs/05_ACCEPTANCE_CRITERIA.md`
8. Reference workbook

The workbook is a baseline prototype, not permission to reduce the target scope.

## Required Codex behavior
Codex must first provide a plan for approval. It must not begin implementation in its first response.
