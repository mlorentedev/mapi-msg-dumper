from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


@dataclass(frozen=True)
class MarkdownEmail:
    received_at: datetime
    subject: str
    sender_name: str
    sender_email: str
    to: str
    cc: str
    entry_id: str
    source_msg_path: Path
    folder_path: str


def render_email_markdown(email: MarkdownEmail, body: str) -> str:
    normalized_body = body.replace("\r\n", "\n").replace("\r", "\n").strip() or "(empty body)"
    safe_subject = _single_line(email.subject) or "no-subject"
    source_msg = email.source_msg_path.as_posix()
    return "\n".join(
        [
            "---",
            'type: "outlook-email"',
            f'received_at: "{email.received_at.isoformat()}"',
            f'subject: "{_escape_yaml(safe_subject)}"',
            f'from_name: "{_escape_yaml(_single_line(email.sender_name))}"',
            f'from_email: "{_escape_yaml(_single_line(email.sender_email))}"',
            f'to: "{_escape_yaml(_single_line(email.to))}"',
            f'cc: "{_escape_yaml(_single_line(email.cc))}"',
            f'entry_id: "{_escape_yaml(_single_line(email.entry_id))}"',
            f'source_msg: "{_escape_yaml(source_msg)}"',
            f'folder: "{_escape_yaml(_single_line(email.folder_path))}"',
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
            f"- **Source MSG**: `{source_msg}`",
            "",
            "## Body",
            "",
            normalized_body,
            "",
        ]
    )


def _single_line(value: str) -> str:
    return " ".join(value.split())


def _escape_yaml(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')
