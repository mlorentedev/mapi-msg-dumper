from datetime import datetime
from pathlib import Path

from mapi_msg_dumper.core.filenames import markdown_file_path, message_file_path, sanitize_subject


def test_sanitize_subject_replaces_invalid_ntfs_chars() -> None:
    assert sanitize_subject('Q1: Sales\\Pipeline/Review*?"<>|') == "Q1_ Sales_Pipeline_Review______"


def test_sanitize_subject_falls_back_to_default() -> None:
    assert sanitize_subject(r'\/:*?"<>|   ') == "no-subject"


def test_sanitize_subject_respects_max_length() -> None:
    assert len(sanitize_subject("x" * 300, max_length=50)) == 50


def test_markdown_and_msg_paths_share_same_stem() -> None:
    received_at = datetime(2026, 2, 23, 8, 11, 40)

    msg_path = message_file_path(Path("exports"), received_at, "Subject A", "entry-123")
    md_path = markdown_file_path(Path("exports/markdown"), received_at, "Subject A", "entry-123")

    assert msg_path.stem == md_path.stem
    assert msg_path.suffix == ".msg"
    assert md_path.suffix == ".md"
