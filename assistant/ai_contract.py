"""Prompt, provider, and response validation for approved R7 assistant calls."""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

ROOT = Path(__file__).resolve().parent
CATALOG = json.loads((ROOT / "prompts" / "catalog.json").read_text(encoding="utf-8"))
PROVIDERS = json.loads((ROOT / "providers.json").read_text(encoding="utf-8"))


class AiContractError(Exception):
    def __init__(self, code: str, user_message: str, technical_message: str, retryable: bool = False):
        super().__init__(technical_message)
        self.code = code
        self.user_message = user_message
        self.retryable = retryable


def invoke(payload: dict[str, Any]) -> dict[str, Any]:
    body = payload.get("payload") if isinstance(payload.get("payload"), dict) else {}
    prompt_id = str(body.get("prompt_id", "")).strip()
    provider_id = str(body.get("provider_id", "")).strip()
    inputs = body.get("inputs")
    if prompt_id not in CATALOG:
        raise AiContractError("unknown_prompt", "The requested AI action is not configured.", f"Unknown prompt_id: {prompt_id}")
    if provider_id not in PROVIDERS:
        raise AiContractError("unknown_provider", "The selected AI provider is not configured.", f"Unknown provider_id: {provider_id}")
    if not isinstance(inputs, dict):
        raise AiContractError("invalid_ai_request", "The AI request inputs are invalid.", "inputs must be an object.")
    prompt = CATALOG[prompt_id]
    missing = [name for name in prompt["required_inputs"] if name not in inputs]
    if missing:
        raise AiContractError("missing_prompt_input", "The AI request is missing required inputs.", "Missing inputs: " + ", ".join(missing))
    structured = validate_output(prompt, obtain_model_output(prompt_id, prompt, inputs, PROVIDERS[provider_id]))
    return {"prompt_id": prompt_id, "provider_id": provider_id, "model": model_name(PROVIDERS[provider_id]), "parameters": {"temperature": prompt["temperature"], "max_output_tokens": prompt["max_output_tokens"]}, "structured_output": structured}


def model_name(provider: dict[str, Any]) -> str:
    return os.environ.get(provider.get("model_environment_variable", ""), provider.get("default_model", ""))


def obtain_model_output(prompt_id: str, prompt: dict[str, Any], inputs: dict[str, Any], provider: dict[str, Any]) -> dict[str, Any]:
    mocked = os.environ.get("CONTENT_ASSISTANT_MOCK_AI_RESPONSE", "")
    if mocked:
        return json.loads(mocked)
    key_name = provider.get("api_key_environment_variable", "")
    api_key = os.environ.get(key_name, "")
    base_url = os.environ.get(provider.get("base_url_environment_variable", ""), provider.get("base_url", ""))
    if not api_key or not base_url or not model_name(provider):
        raise AiContractError("provider_credentials_unavailable", "The selected provider is not ready. Configure its environment variables and try again.", "Provider base URL, model, or API key environment variable is unavailable.")
    system = "Return only valid JSON. Follow the requested output schema. Cite supplied Source or Evidence IDs only; never invent citations. AI output is a pending suggestion and must not state it has been approved."
    user = json.dumps({"prompt_id": prompt_id, "instruction": prompt["instruction"], "inputs": inputs, "required_output_schema": prompt["output_schema"]}, ensure_ascii=False)
    request_body = json.dumps({"model": model_name(provider), "messages": [{"role": "system", "content": system}, {"role": "user", "content": user}], "temperature": prompt["temperature"], "max_tokens": prompt["max_output_tokens"], "response_format": {"type": "json_object"}}).encode("utf-8")
    request = Request(base_url, data=request_body, headers={"Authorization": "Bearer " + api_key, "Content-Type": "application/json"}, method="POST")
    try:
        with urlopen(request, timeout=45) as response:
            result = json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        raise AiContractError("provider_error", "The AI provider returned an error. Try again later or check the selected model.", f"Provider HTTP {exc.code}", exc.code >= 500) from exc
    except URLError as exc:
        raise AiContractError("provider_unavailable", "The AI provider could not be reached. Check the network and try again.", repr(exc.reason), True) from exc
    try:
        return json.loads(result["choices"][0]["message"]["content"])
    except (KeyError, IndexError, TypeError, json.JSONDecodeError) as exc:
        raise AiContractError("invalid_provider_response", "The AI provider returned an invalid structured response.", repr(exc)) from exc


def validate_output(prompt: dict[str, Any], output: Any) -> dict[str, Any]:
    if not isinstance(output, dict):
        raise AiContractError("invalid_ai_response", "The AI response is not a JSON object.", "Structured output is not an object.")
    for name, expected_type in prompt["output_schema"].items():
        value = output.get(name)
        if expected_type == "string" and not isinstance(value, str):
            raise AiContractError("invalid_ai_response", "The AI response is missing required text.", f"Expected string: {name}")
        if expected_type == "array" and not isinstance(value, list):
            raise AiContractError("invalid_ai_response", "The AI response is missing a required list.", f"Expected array: {name}")
        if expected_type == "number" and (not isinstance(value, (int, float)) or isinstance(value, bool)):
            raise AiContractError("invalid_ai_response", "The AI response is missing a required number.", f"Expected number: {name}")
    return output
