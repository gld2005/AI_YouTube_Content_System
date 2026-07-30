import json
import os
import threading
import tempfile
import unittest
from http.client import HTTPConnection
from http.server import ThreadingHTTPServer
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "assistant"))
from server import AssistantHandler


class ServerTests(unittest.TestCase):
    def setUp(self):
        self.previous_mock_response = os.environ.get("CONTENT_ASSISTANT_MOCK_AI_RESPONSE")
        self.server = ThreadingHTTPServer(("127.0.0.1", 0), AssistantHandler)
        self.server.token = "test-token"
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()

    def tearDown(self):
        self.server.shutdown()
        self.thread.join()
        self.server.server_close()
        if self.previous_mock_response is None:
            os.environ.pop("CONTENT_ASSISTANT_MOCK_AI_RESPONSE", None)
        else:
            os.environ["CONTENT_ASSISTANT_MOCK_AI_RESPONSE"] = self.previous_mock_response

    def request(self, method, path, body=None, token="test-token"):
        connection = HTTPConnection("127.0.0.1", self.server.server_port)
        headers = {"X-Local-Assistant-Token": token}
        if body is not None:
            headers["Content-Type"] = "application/json"
        connection.request(method, path, body=json.dumps(body) if body is not None else None, headers=headers)
        response = connection.getresponse()
        try:
            return response.status, json.loads(response.read())
        finally:
            connection.close()

    def test_health_is_available_without_token(self):
        status, body = self.request("GET", "/v1/health", token="")
        self.assertEqual(status, 200)
        self.assertEqual(body["status"], "healthy")

    def test_capabilities_requires_token(self):
        status, body = self.request("GET", "/v1/capabilities", token="bad")
        self.assertEqual(status, 401)
        self.assertEqual(body["error"]["code"], "unauthorized")

    def test_folder_creation(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory) / "Project"
            status, body = self.request("POST", "/v1/folders/create", {"request_id": "test", "payload": {"project_root": str(root), "folders": ["Planning", "Research"]}})
            self.assertEqual(status, 200)
            self.assertEqual(body["status"], "completed")
            self.assertTrue((root / "Planning").is_dir())

    def test_folder_creation_rejects_path_traversal(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            status, body = self.request("POST", "/v1/folders/create", {"payload": {"project_root": temporary_directory, "folders": ["../outside"]}})
            self.assertEqual(status, 400)
            self.assertEqual(body["error"]["code"], "invalid_folder_name")

    def test_ai_candidate_is_pending_approval(self):
        os.environ["CONTENT_ASSISTANT_MOCK_AI_RESPONSE"] = json.dumps({"candidate": "Polished text.", "changes": [], "citations": [], "duration_seconds": 5})
        status, body = self.request("POST", "/v1/ai/invoke", {"request_id": "ai-test", "payload": {"provider_id": "openai", "prompt_id": "polish_paragraph", "inputs": {"paragraph": "Original text.", "type": "Narration", "status": "Drafting", "constraints": "Keep it factual."}}})
        self.assertEqual(status, 200)
        self.assertEqual(body["status"], "pending_approval")
        self.assertEqual(body["data"]["structured_output"]["candidate"], "Polished text.")

    def test_ai_rejects_missing_required_input(self):
        status, body = self.request("POST", "/v1/ai/invoke", {"payload": {"provider_id": "openai", "prompt_id": "polish_paragraph", "inputs": {"paragraph": "Original text."}}})
        self.assertEqual(status, 400)
        self.assertEqual(body["error"]["code"], "missing_prompt_input")
