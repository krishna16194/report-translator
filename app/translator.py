"""Translate the content model using deep-translator's free Google backend.

GoogleTranslator (the public endpoint) caps each request at ~5000 characters,
so long text is split into chunks on sentence/word boundaries, translated, and
re-joined. Every translated string is cached for the duration of a request to
avoid re-translating repeated cells (common in spreadsheets).
"""
from __future__ import annotations

from typing import Any

from deep_translator import GoogleTranslator

# Languages offered in the UI. Codes are what GoogleTranslator expects.
LANGUAGES: dict[str, str] = {
    "en": "English",
    "es": "Spanish",
    "fr": "French",
    "de": "German",
    "it": "Italian",
    "pt": "Portuguese",
    "nl": "Dutch",
    "ru": "Russian",
    "ar": "Arabic",
    "hi": "Hindi",
    "ta": "Tamil",
    "te": "Telugu",
    "kn": "Kannada",
    "ml": "Malayalam",
    "bn": "Bengali",
    "ja": "Japanese",
    "ko": "Korean",
    "zh-CN": "Chinese (Simplified)",
    "zh-TW": "Chinese (Traditional)",
}

_MAX_CHARS = 4800


class Translator:
    def __init__(self, target_lang: str, source_lang: str = "auto") -> None:
        if target_lang not in LANGUAGES:
            raise ValueError(f"Unsupported target language: {target_lang}")
        self._engine = GoogleTranslator(source=source_lang, target=target_lang)
        self._cache: dict[str, str] = {}

    def text(self, value: str) -> str:
        """Translate a single string, using the per-instance cache."""
        if not value or not value.strip():
            return value
        if value in self._cache:
            return self._cache[value]

        if len(value) <= _MAX_CHARS:
            result = self._engine.translate(value)
        else:
            result = " ".join(self._engine.translate(c) for c in _chunk(value))

        result = result if result is not None else value
        self._cache[value] = result
        return result

    def blocks(self, blocks: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """Translate every block, preserving structure."""
        out: list[dict[str, Any]] = []
        for block in blocks:
            btype = block["type"]
            if btype in ("paragraph", "heading"):
                new = dict(block)
                new["text"] = self.text(block["text"])
                out.append(new)
            elif btype == "table":
                new = dict(block)
                new["rows"] = [[self.text(c) for c in row] for row in block["rows"]]
                out.append(new)
            else:
                out.append(block)
        return out


def _chunk(text: str, limit: int = _MAX_CHARS) -> list[str]:
    """Split text into <=limit pieces, preferring sentence/word boundaries."""
    chunks: list[str] = []
    remaining = text
    while len(remaining) > limit:
        window = remaining[:limit]
        # Prefer to break after a sentence end, then a space.
        split_at = max(window.rfind(". "), window.rfind("\n"))
        if split_at < limit // 2:
            split_at = window.rfind(" ")
        if split_at <= 0:
            split_at = limit
        chunks.append(remaining[: split_at + 1])
        remaining = remaining[split_at + 1 :]
    if remaining:
        chunks.append(remaining)
    return chunks
