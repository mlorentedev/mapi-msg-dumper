from datetime import datetime
from pathlib import Path

from mapi_msg_dumper.core.markdown import ExtractedEmail, render_email_markdown


def test_render_email_markdown_contains_metadata_and_body() -> None:
    email = ExtractedEmail(
        received_at=datetime(2026, 2, 23, 8, 11, 40),
        subject='DUO "Authentication" Issue',
        sender_name="IT Support",
        sender_email="it-support@example.com",
        to="manu@example.com",
        cc="team@example.com",
        entry_id="abc123",
        folder_path="Inbox",
        body_text="First line\r\nSecond line",
        tags=["support", "auth"],
    )
    content = render_email_markdown(email, source_raw_path=Path(r"C:/exports/2026/02/mail.msg"))

    assert 'type: "email"' in content
    assert 'subject: "DUO \\"Authentication\\" Issue"' in content
    assert 'tags:\n  - "support"\n  - "auth"' in content
    assert "## Body" in content
    assert "First line\nSecond line" in content
