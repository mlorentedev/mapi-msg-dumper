from __future__ import annotations

import hashlib
import re
from datetime import datetime
from pathlib import Path

_INVALID_FILENAME_CHARS = re.compile(r"[\\/:*?\"<>|]")
_WHITESPACE = re.compile(r"\s+")


def sanitize_subject(subject: str, max_length: int = 90) -> str:
    cleaned = _INVALID_FILENAME_CHARS.sub("_", subject)
    cleaned = _WHITESPACE.sub(" ", cleaned).strip()
    cleaned = cleaned.rstrip(". ")
    if not cleaned or cleaned.replace("_", "").strip() == "":
        cleaned = "no-subject"
    return cleaned[:max_length]


def stable_message_stem(received_at: datetime, subject: str, entry_id: str) -> str:
    stable_suffix = hashlib.sha1(_safe_entry_id(entry_id).encode("utf-8")).hexdigest()[:10]
    return f"{received_at:%Y%m%d-%H%M%S}_{sanitize_subject(subject)}_{stable_suffix}"


def message_file_path(output_root: Path, received_at: datetime, subject: str, entry_id: str) -> Path:
    target_dir = output_root / f"{received_at:%Y}" / f"{received_at:%m}"
    filename = f"{stable_message_stem(received_at, subject, entry_id)}.msg"
    return target_dir / filename


def markdown_file_path(output_root: Path, received_at: datetime, subject: str, entry_id: str) -> Path:
    target_dir = output_root / f"{received_at:%Y}" / f"{received_at:%m}"
    filename = f"{stable_message_stem(received_at, subject, entry_id)}.md"
    return target_dir / filename


def _safe_entry_id(entry_id: str) -> str:
    return entry_id.strip() or "missing-entry-id"
