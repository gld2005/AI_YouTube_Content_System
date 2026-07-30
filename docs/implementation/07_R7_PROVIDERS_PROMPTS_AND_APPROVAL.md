# R7 Providers, Prompts, and Approval

## Objective

R7 enables provider-configured, structured AI requests while preserving the rule that no AI output can change workbook source-of-truth content without explicit user approval.

## Components

| Component | Behavior |
|---|---|
| `assistant/providers.json` | OpenAI, DeepSeek, OpenRouter, Alibaba-compatible, and custom OpenAI-compatible non-secret provider presets. |
| `assistant/prompts/catalog.json` | External English prompt definitions for all approved R7 AI functions, required inputs, output schemas, temperature, and output limits. |
| `assistant/ai_contract.py` | Validates prompt inputs and structured output, reads credentials only from environment variables, and calls OpenAI-compatible endpoints. |
| `assistant/server.py` | Enables `POST /v1/ai/invoke` and always labels successful output `pending_approval`. |
| `src/vba/modAiAssistant.bas` | Requests an AI candidate for the selected paragraph and queues it in the existing Approval Center without changing the script. |

## Provider and Secret Rules

Provider URLs, model selection, and API-key environment-variable names are source controlled. Secrets are never present in the workbook, repository configuration, VBA, logs, or request payloads. The assistant reads the selected provider key from the Windows process environment only. The custom compatible provider reads its base URL, model, and key from environment variables.

## Approval Rules

The assistant returns validated JSON in a `pending_approval` response. Excel can queue an expand, polish, compress, or rewrite candidate only through `QueueAiSuggestion`; that queue does not change the script. Existing approval commands remain the only way to adopt or reject it. Analysis and packaging prompt outputs have no adoption target.

## Prompt Coverage

The catalog contains chapter outline, hook, four paragraph rewrite actions, five quality-analysis actions, visual ideas, asset list, duration estimation, research questions, and five-concept packaging. Each prompt instructs the provider to return JSON only, cite supplied source or evidence IDs only, and never claim a suggestion is approved.

## R7 API Behavior

`POST /v1/ai/invoke` accepts `provider_id`, `prompt_id`, and `inputs` under the common request payload. It returns validated `structured_output`, selected model, and prompt parameters. Missing inputs, missing credentials, provider failures, timeouts, and malformed model JSON return the standard structured error object. The assistant uses a 45-second provider timeout and does not retry a provider request automatically.

## Verification Evidence

| Check | Result |
|---|---|
| Assistant unit suite | PASS: six tests cover health, authorization, folder safety, pending AI output, and prompt-input validation. |
| R7 workbook generation | PASS: `working/AI_YouTube_Content_System_R7_providers_prompts_final.xlsm` was generated from the protected R2 source plus current VBA modules. |
| VBA regression | PASS: `RunR5SmokeTest` executed in the generated R7 workbook. |
| R7 executable build | PASS: `assistant/dist/ContentAssistant.exe` was rebuilt including prompt and provider resource files. |
| Executable AI request | PASS: the executable returned `pending_approval` for a structured mock-provider response. |
