# AGENTS.md — Mandatory Operating Rules

## Mission
Build the complete AI YouTube Content Planning and Production Workbook described in this repository. The result is a content-first planning system, not merely a task tracker.

## Mandatory first response from Codex
Before changing any file, Codex must:

1. Read every Markdown file in this repository and inspect the reference workbook.
2. Produce its own implementation plan, architecture proposal, risk assessment, validation strategy, and expected file changes.
3. Map the proposed work to the requirement IDs and acceptance criteria in the documentation.
4. Wait for explicit user approval before implementing anything.

The repository intentionally does not contain an implementation roadmap. Codex owns the planning process.

## Scope protection
- Every confirmed requirement in this repository is in scope.
- Codex may reorganize architecture, implementation order, internal schemas, or technical mechanisms.
- Codex must not silently remove, merge away, weaken, postpone, or reinterpret a confirmed capability.
- If a requirement is technically unsafe, infeasible, or incompatible with Excel, Codex must propose an equivalent alternative and wait for approval.
- Phased delivery is allowed, but the final approved target must still cover the complete specification.

## Change safety
- Never overwrite the only copy of the reference workbook.
- Develop on a working copy and preserve migration paths for existing data.
- Back up workbooks before structural migrations.
- Keep formulas, IDs, links, VBA modules, and local-assistant contracts testable and documented.

## Language and locale
- All developer documentation, code comments, VBA comments, Python comments, configuration names, log messages, tests, and technical error messages must be written in English.
- The user-facing Excel interface should remain in Simplified Chinese unless the user explicitly requests a locale change.
- Use Microsoft YaHei for Chinese body text and Aptos for English text, numbers, and metrics where practical.

## Approval and automation safety
- AI output must never modify approved workbook content without explicit user approval.
- Project-state transitions must be suggestions until the user confirms them.
- New AI-generated structures and template changes must require approval.
- API keys must never be stored as plaintext in the workbook.
- Prefer environment variables; use session-only input as a fallback.

## Primary compatibility target
- Windows desktop Excel with macros enabled.
- The workbook may be viewable elsewhere, but VBA-dependent actions are not required to run in Excel for the web or mobile Excel.

## Definition of success
The final system must satisfy the acceptance criteria, preserve the confirmed design language, and feel like a coherent creator content workspace rather than a collection of unrelated spreadsheets.
