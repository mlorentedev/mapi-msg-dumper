from datetime import datetime
from pathlib import Path

from mapi_msg_dumper.core.markdown import MarkdownEmail, render_email_markdown


def test_render_email_markdown_contains_metadata_and_body() -> None:
    content = render_email_markdown(
        MarkdownEmail(
            received_at=datetime(2026, 2, 23, 8, 11, 40),
            subject='DUO "Authentication" Issue',
            sender_name="IT Support",
            sender_email="it-support@example.com",
            to="manu@example.com",
            cc="team@example.com",
            entry_id="abc123",
            source_msg_path=Path(r"C:\exports\2026\02\mail.msg"),
            folder_path="Inbox",
        ),
        body="First line\r\nSecond line",
    )

    assert 'type: "outlook-email"' in content
    assert 'subject: "DUO \\"Authentication\\" Issue"' in content
    assert "## Body" in content
    assert "First line\nSecond line" in content
