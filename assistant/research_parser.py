"""Local research parsing with explicit recovery messages."""

from __future__ import annotations

import html
import io
import re
from pathlib import Path
from typing import Any
from urllib.error import URLError
from urllib.request import Request, urlopen

from docx import Document
from pypdf import PdfReader
from youtube_transcript_api import YouTubeTranscriptApi


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
    if kind == "YouTube":
        return parse_youtube(url)
    try:
        request = Request(url, headers={"User-Agent": "AI-YouTube-Content-System/0.1"})
        with urlopen(request, timeout=30) as response:
            content_type = response.headers.get_content_type()
            data = response.read(1_500_000)
    except URLError as exc:
        raise ParseError("source_unavailable", "The source could not be reached. Check the link, network, or access permissions.", repr(exc.reason), ["Retry later", "Paste source text", "Create a manual source record"]) from exc
    if content_type == "application/pdf":
        return parse_pdf_bytes(data, url)
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
        return parse_pdf_bytes(path.read_bytes(), path.name)
    if suffix == ".docx":
        try:
            document = Document(path)
            text = "\n".join(paragraph.text for paragraph in document.paragraphs).strip()
        except Exception as exc:
            raise ParseError("document_parse_failed", "The DOCX file could not be read. Export text or paste the relevant passage.", repr(exc), ["Export text", "Paste text", "Create a manual source record"]) from exc
        return {"source_type": "Local Document", "title": path.name, "text": text[:12000], "warnings": ["Review imported text before using it as evidence."]}
    raise ParseError("unsupported_file_type", "This local file type is not supported for automatic extraction.", f"Unsupported extension: {suffix}", ["Export text or transcript", "Paste text", "Create a manual source record"])


def parse_pdf_bytes(data: bytes, title: str) -> dict[str, Any]:
    try:
        reader = PdfReader(io.BytesIO(data))
        text = "\n".join(page.extract_text() or "" for page in reader.pages).strip()
    except Exception as exc:
        raise ParseError("pdf_parse_failed", "The PDF could not be read. It may be encrypted or image-only; use OCR or paste text.", repr(exc), ["Use OCR for scanned PDFs", "Export PDF text", "Create a manual source record"]) from exc
    if not text:
        raise ParseError("pdf_has_no_text_layer", "This PDF has no readable text layer. Use OCR or paste text.", "No text was extracted from PDF.", ["Use OCR for scanned PDFs", "Paste text", "Create a manual source record"])
    return {"source_type": "PDF", "title": title, "text": text[:12000], "warnings": ["Review imported text before creating evidence."]}


def parse_youtube(url: str) -> dict[str, Any]:
    match = re.search(r"(?:v=|youtu\.be/|shorts/)([A-Za-z0-9_-]{11})", url)
    if not match:
        raise ParseError("invalid_youtube_url", "The YouTube URL does not contain a valid video ID.", url, ["Paste a full YouTube video URL", "Import a subtitle file", "Create a manual source record"])
    try:
        transcript = YouTubeTranscriptApi().fetch(match.group(1))
        text = "\n".join(snippet.text for snippet in transcript).strip()
    except Exception as exc:
        raise ParseError("youtube_transcript_unavailable", "YouTube subtitles are unavailable or blocked. Import subtitles, paste the transcript, or create a manual source record.", repr(exc), ["Import SRT or VTT subtitles", "Paste transcript text", "Create a manual source record"]) from exc
    return {"source_type": "YouTube", "title": "YouTube " + match.group(1), "text": text[:12000], "warnings": ["Review subtitles before creating evidence."]}
