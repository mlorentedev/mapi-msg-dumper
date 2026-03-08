from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


@dataclass
class ExtractedEmail:
    received_at: datetime
    subject: str
    sender_name: str
    sender_email: str
    to: str
    cc: str
    entry_id: str
    folder_path: str
    body_text: str
    tags: list[str]


def render_email_markdown(email: ExtractedEmail, source_raw_path: Path | None) -> str:
    normalized_body = email.body_text.replace("\r\n", "\n").replace("\r", "\n").strip() or "(empty body)"
    safe_subject = _single_line(email.subject) or "no-subject"
    source_raw = source_raw_path.as_posix() if source_raw_path else ""

    tags_yaml = "\n".join(f'  - "{_escape_yaml(t)}"' for t in email.tags) if email.tags else "[]"

    yaml_lines = [
        "---",
        'type: "email"',
        f'received_at: "{email.received_at.isoformat()}"',
        f'subject: "{_escape_yaml(safe_subject)}"',
        f'from_name: "{_escape_yaml(_single_line(email.sender_name))}"',
        f'from_email: "{_escape_yaml(_single_line(email.sender_email))}"',
        f'to: "{_escape_yaml(_single_line(email.to))}"',
        f'cc: "{_escape_yaml(_single_line(email.cc))}"',
        f'entry_id: "{_escape_yaml(_single_line(email.entry_id))}"',
        f'source_raw: "{_escape_yaml(source_raw)}"',
        f'folder: "{_escape_yaml(_single_line(email.folder_path))}"',
        f"tags:\n{tags_yaml}" if email.tags else f"tags: {tags_yaml}",
        "---",
        "",
        f"# {safe_subject}",
        "",
        "## Metadata",
        f"- **Received**: {email.received_at.isoformat()}",
        f"- **From**: {_single_line(email.sender_name)} <{_single_line(email.sender_email)}>",
        f"- **To**: {_single_line(email.to)}",
        f"- **CC**: {_single_line(email.cc)}",
        f"- **Folder**: {_single_line(email.folder_path)}",
        f"- **Source Raw**: `{source_raw}`",
        "",
        "## Body",
        "",
        normalized_body,
        "",
    ]
    return "\n".join(yaml_lines)


def _single_line(value: str) -> str:
    return " ".join(value.split())


def _escape_yaml(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')
