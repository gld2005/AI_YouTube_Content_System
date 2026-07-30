"""Local-first HTTP assistant for the AI YouTube Content System."""

from __future__ import annotations

import json
import os
import secrets
import sys
import uuid
from datetime import datetime, timezone
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any

from ai_contract import AiContractError, invoke as invoke_ai

VERSION = "0.1.0"
DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 8765


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def error(code: str, user_message: str, technical_message: str, retryable: bool = False, recovery_actions: list[str] | None = None) -> dict[str, Any]:
    return {
        "code": code,
        "user_message": user_message,
        "technical_message": technical_message,
        "retryable": retryable,
        "recovery_actions": recovery_actions or [],
    }


class AssistantHandler(BaseHTTPRequestHandler):
    server_version = "ContentAssistant/0.1"

    def log_message(self, _format: str, *_args: Any) -> None:
        return

    @property
    def token(self) -> str:
        return self.server.token  # type: ignore[attr-defined]

    def send_json(self, status: HTTPStatus, body: dict[str, Any]) -> None:
        encoded = json.dumps(body, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def authorized(self) -> bool:
        return secrets.compare_digest(self.headers.get("X-Local-Assistant-Token", ""), self.token)

    def response(self, request_id: str, status: str, data: dict[str, Any] | None = None, warnings: list[str] | None = None, issue: dict[str, Any] | None = None) -> dict[str, Any]:
        return {"request_id": request_id, "operation_id": str(uuid.uuid4()), "status": status, "data": data or {}, "warnings": warnings or [], "error": issue}

    def read_request(self) -> tuple[str, dict[str, Any] | None]:
        try:
            length = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            request_id = str(payload.get("request_id") or uuid.uuid4())
            return request_id, payload
        except (ValueError, json.JSONDecodeError):
            return str(uuid.uuid4()), None

    def do_GET(self) -> None:
        if self.path == "/v1/health":
            self.send_json(HTTPStatus.OK, self.response(str(uuid.uuid4()), "healthy", {"version": VERSION, "time": utc_now(), "mode": "local"}))
            return
        if not self.authorized():
            self.send_json(HTTPStatus.UNAUTHORIZED, self.response(str(uuid.uuid4()), "error", issue=error("unauthorized", "The local assistant token is invalid.", "Missing or invalid X-Local-Assistant-Token.")))
            return
        if self.path == "/v1/capabilities":
            self.send_json(HTTPStatus.OK, self.response(str(uuid.uuid4()), "ok", {
                "providers": [],
                "parsers": [],
                "actions": ["folders.create", "ai.invoke"],
                "endpoints": ["/v1/health", "/v1/capabilities", "/v1/ai/invoke", "/v1/folders/create", "/v1/operations/{id}/cancel"],
                "cloud_compatible": True,
            }))
            return
        self.send_json(HTTPStatus.NOT_FOUND, self.response(str(uuid.uuid4()), "error", issue=error("not_found", "The requested local assistant endpoint was not found.", self.path)))

    def do_POST(self) -> None:
        request_id, payload = self.read_request()
        if payload is None:
            self.send_json(HTTPStatus.BAD_REQUEST, self.response(request_id, "error", issue=error("invalid_json", "The assistant request is not valid JSON.", "JSON parsing failed.")))
            return
        if not self.authorized():
            self.send_json(HTTPStatus.UNAUTHORIZED, self.response(request_id, "error", issue=error("unauthorized", "The local assistant token is invalid.", "Missing or invalid X-Local-Assistant-Token.")))
            return
        if self.path == "/v1/folders/create":
            self.create_folders(request_id, payload)
            return
        if self.path == "/v1/ai/invoke":
            self.invoke_ai(request_id, payload)
            return
        if self.path.startswith("/v1/operations/") and self.path.endswith("/cancel"):
            self.send_json(HTTPStatus.NOT_IMPLEMENTED, self.response(request_id, "error", issue=error("not_enabled", "Operation cancellation is not enabled yet.", "R6 has no long-running operations to cancel.", False, ["Wait for the current request to finish"])))
            return
        self.send_json(HTTPStatus.NOT_IMPLEMENTED, self.response(request_id, "error", issue=error("not_enabled", "This assistant action is not enabled yet.", f"The current assistant does not implement {self.path}.", False, ["Check capabilities", "Use the workbook manual workflow"])))

    def invoke_ai(self, request_id: str, payload: dict[str, Any]) -> None:
        try:
            data = invoke_ai(payload)
        except AiContractError as exc:
            status = HTTPStatus.SERVICE_UNAVAILABLE if exc.retryable else HTTPStatus.BAD_REQUEST
            self.send_json(status, self.response(request_id, "error", issue=error(exc.code, exc.user_message, str(exc), exc.retryable, ["Check provider settings", "Retry only after correcting the request or service"])))
            return
        self.send_json(HTTPStatus.OK, self.response(request_id, "pending_approval", data, ["AI output is pending approval and has not modified workbook content."]))

    def create_folders(self, request_id: str, payload: dict[str, Any]) -> None:
        body = payload.get("payload") if isinstance(payload.get("payload"), dict) else {}
        root_text = str(body.get("project_root", "")).strip()
        folder_names = body.get("folders")
        if not root_text or not isinstance(folder_names, list) or not all(isinstance(item, str) and item.strip() for item in folder_names):
            self.send_json(HTTPStatus.BAD_REQUEST, self.response(request_id, "error", issue=error("invalid_folder_request", "Choose a valid project root and folder template.", "project_root or folders is invalid.")))
            return
        root = Path(root_text).expanduser()
        if not root.is_absolute():
            self.send_json(HTTPStatus.BAD_REQUEST, self.response(request_id, "error", issue=error("invalid_project_root", "Choose an absolute project root path.", "project_root must be absolute.")))
            return
        for name in folder_names:
            path = Path(name)
            if path.is_absolute() or ".." in path.parts:
                self.send_json(HTTPStatus.BAD_REQUEST, self.response(request_id, "error", issue=error("invalid_folder_name", "The folder template contains an unsafe path.", "Folder names cannot be absolute or traverse to a parent folder.")))
                return
        created: list[str] = []
        try:
            root.mkdir(parents=True, exist_ok=True)
            for name in folder_names:
                target = root / name
                target.mkdir(parents=True, exist_ok=True)
                created.append(str(target))
        except OSError as exc:
            self.send_json(HTTPStatus.FORBIDDEN, self.response(request_id, "error", {"created_paths": created}, issue=error("folder_creation_failed", "The project folders could not be created. Check the path and permissions.", repr(exc), False, ["Choose another project root", "Check folder permissions"])))
            return
        self.send_json(HTTPStatus.OK, self.response(request_id, "completed", {"created_paths": created, "rollback": "No rollback required; existing folders were preserved."}))


def main() -> None:
    host = os.environ.get("CONTENT_ASSISTANT_HOST", DEFAULT_HOST)
    port = int(os.environ.get("CONTENT_ASSISTANT_PORT", str(DEFAULT_PORT)))
    token = os.environ.get("CONTENT_ASSISTANT_TOKEN", "")
    if not token:
        print("CONTENT_ASSISTANT_TOKEN must be set for local assistant startup.", file=sys.stderr)
        raise SystemExit(2)
    server = ThreadingHTTPServer((host, port), AssistantHandler)
    server.token = token  # type: ignore[attr-defined]
    print(json.dumps({"host": host, "port": port, "version": VERSION}), flush=True)
    server.serve_forever()


if __name__ == "__main__":
    main()
