# Phase 8: Approval and Publish Review

Phase 8 introduces an approval-gated AI suggestion record and a safe publish-review record action.

## Delivered behavior

- `QueueAiSuggestion` writes a complete Pending record to the AI Approval Center without modifying the script.
- `ApproveFullAdoption` is an explicit action. It applies the pending candidate only to the linked paragraph, records Full Adoption, and adds a compact script-page audit line.
- `RejectAiSuggestion` retains the rejection reason without modifying source text.
- `AddPublishReview` creates a publish-review record for a video and platform.
- No API key, provider call, or model output transport is included in this phase; queued candidates may be created by a future local assistant or entered manually.

## Validation

`RunPhase8SmokeTest` proves that a queued candidate leaves approved text unchanged, then verifies that explicit approval changes only the linked paragraph and creates a publish-review record.

`PASS: AI suggestion approval is gated and publish review was recorded.`

## Requirement traceability

This phase advances AC-134, AC-136, AC-137, AC-138, AC-139, and the workbook-side record model for AC-130 through AC-133. Provider configuration, session-only secrets, and the local assistant remain open.
