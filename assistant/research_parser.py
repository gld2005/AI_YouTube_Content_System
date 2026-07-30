"""Local research parsing with explicit recovery messages."""

from __future__ import annotations

import html
import re
from pathlib import Path
from typing import Any
from urllib.error import URLError
from urllib.request import Request, urlopen


class ParseError(Exception):
    def __init__(self, code: str, user_message: str, technical_message: str, recovery_actions: list[str]):
        super().__init__(technical_message)
        self.code = code
        self.user_message = user_message
        self.recovery_actions = recovery_actions


def source_type(location: str) -> str:
    lowered = location.lower()
    if "youtube.com" in lowered or "youtu.be" in lowered: return "YouTube"
    if "bilibili.com" in lowered: return "Bilibili"
    if "reddit.com" in lowered: return "Reddit"
    if "x.com" in lowered or "twitter.com" in lowered: return "X / Twitter"
    return "Web"


def parse_url(url: str) -> dict[str, Any]:
    kind = source_type(url)
    if kind in {"Bilibili", "Reddit", "X / Twitter"}:
        raise ParseError("semi_automatic_source", "This platform may block automated extraction. Paste text, subtitles, or an exported file to continue.", f"Semi-automatic platform: {kind}", ["Paste source text", "Import subtitles or an export", "Create a manual source record"])
    try:
        request = Request(url, headers={"User-Agent": "AI-YouTube-Content-System/0.1"})
        with urlopen(request, timeout=30) as response:
            content_type = response.headers.get_content_type()
            data = response.read(1_500_000)
    except URLError as exc:
        raise ParseError("source_unavailable", "The source could not be reached. Check the link, network, or access permissions.", repr(exc.reason), ["Retry later", "Paste source text", "Create a manual source record"]) from exc
    if content_type == "application/pdf":
        raise ParseError("pdf_parser_unavailable", "The PDF was found, but local PDF text extraction is not enabled. Import exported text or add a manual source record.", "PDF extraction requires a dedicated parser.", ["Export PDF text", "Paste text", "Create a manual source record"])
    text = html.unescape(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", data.decode("utf-8", errors="replace")))).strip()
    title_match = re.search(r"<title[^>]*>(.*?)</title>", data.decode("utf-8", errors="replace"), re.I | re.S)
    title = html.unescape(re.sub(r"\s+", " ", title_match.group(1))).strip() if title_match else url
    return {"source_type": kind, "title": title[:300], "text": text[:12000], "warnings": ["Parsed metadata and text are not AI-generated evidence."]}


def parse_file(path_text: str) -> dict[str, Any]:
    path = Path(path_text)
    if not path.is_absolute() or not path.is_file():
        raise ParseError("invalid_local_file", "Choose an existing absolute local file path.", f"Invalid path: {path_text}", ["Choose another file", "Paste text", "Create a manual source record"])
    suffix = path.suffix.lower()
    if suffix in {".txt", ".md", ".csv", ".html", ".htm", ".srt", ".vtt", ".json"}:
        text = path.read_text(encoding="utf-8", errors="replace")
        return {"source_type": "Local Document" if suffix not in {".srt", ".vtt"} else "Transcript", "title": path.name, "text": text[:12000], "warnings": ["Review imported text before using it as evidence."]}
    if suffix == ".pdf":
        raise ParseError("pdf_parser_unavailable", "Local PDF text extraction is not enabled. Import exported text or paste the relevant passage.", "PDF extraction requires a dedicated parser.", ["Export PDF text", "Paste text", "Create a manual source record"])
    raise ParseError("unsupported_file_type", "This local file type is not supported for automatic extraction.", f"Unsupported extension: {suffix}", ["Export text or transcript", "Paste text", "Create a manual source record"])
