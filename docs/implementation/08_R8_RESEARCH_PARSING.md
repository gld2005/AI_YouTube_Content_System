# R8 Research Parsing and Recovery

## Objective

R8 adds local source parsing with visible recovery paths. Parsed metadata and extracted text may be reviewed as source material, but AI summaries, evidence cards, risk labels, and claim links remain approval-gated.

## Supported Paths

| Input | R8 behavior |
|---|---|
| Ordinary web pages and news articles | Attempts title and readable text extraction. |
| YouTube | Attempts direct page extraction; unavailable subtitles require a pasted subtitle or export. |
| Bilibili, Reddit, X / Twitter | Returns a clear semi-automatic recovery message because access may be blocked. |
| URL PDF and local PDF | Returns a PDF-parser recovery message; it does not falsely claim extraction. |
| Local text, Markdown, CSV, HTML, JSON, SRT, VTT | Extracts local text only from an explicitly provided absolute path. |
| Podcasts and other audio | Requests a transcript or exported text. |

## Endpoints

- `POST /v1/sources/parse` accepts `payload.url`.
- `POST /v1/files/parse` accepts `payload.path`.

Both return the standard response schema. Failures include a stable error code, user message, technical message, and manual recovery actions. No parser silently creates evidence or alters approved content.

`ImportResearchUrl` is the Excel entry action. It confirms parsing before recording only source metadata through the existing deduplication workflow. It does not create AI summaries or evidence cards.

## Verification Evidence

| Check | Result |
|---|---|
| Assistant parser suite | PASS: transcript parsing and blocked-platform manual recovery are covered alongside R6/R7 tests. |
| R8 workbook and VBA regression | PASS: R8 workbook was generated and `RunR5SmokeTest` passed. |
| R8 executable | PASS: rebuilt executable parsed a local VTT transcript over loopback. |
