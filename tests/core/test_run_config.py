import json
from pathlib import Path

import pytest

from mapi_msg_dumper.core.folders_config import FolderNode
from mapi_msg_dumper.core.run_config import load_run_config


def test_load_run_config_defaults(tmp_path: Path) -> None:
    config = tmp_path / "run.json"
    config.write_text("{}", encoding="utf-8")

    parsed = load_run_config(config)

    assert parsed.provider.type == "outlook"
    assert parsed.provider.config == {}

    assert parsed.folders == [FolderNode("Inbox", [])]

    assert parsed.outputs.raw_root == tmp_path / "exports"
    assert parsed.outputs.save_raw is True
    assert parsed.outputs.checkpoint_file is None
    assert parsed.outputs.markdown_root is None

    assert parsed.extraction.cadence == "monthly"
    assert parsed.extraction.start_date is None
    assert parsed.extraction.end_date is None
    assert parsed.extraction.manual is False
    assert parsed.extraction.max_windows is None
    assert parsed.extraction.dry_run is False

    assert parsed.verbose is False


def test_load_run_config_with_folders_and_all_options(tmp_path: Path) -> None:
    config = tmp_path / "run.json"
    config.write_text(
        json.dumps(
            {
                "provider": {
                    "type": "thunderbird",
                    "config": {"profile_path": "/some/path"}
                },
                "folders": [{"path": r"Inbox\Finance", "tags": ["finance"]}, r"Inbox\HR"],
                "extraction": {
                    "cadence": "biweekly",
                    "manual": True,
                    "start_date": "2020-01-01",
                    "end_date": "2020-01-31",
                    "max_windows": 2,
                    "dry_run": True
                },
                "outputs": {
                    "raw_root": "./exports/raw",
                    "checkpoint_file": "./exports/checkpoints/folder.json",
                    "markdown_root": "./exports/markdown",
                    "save_raw": False
                },
                "verbose": True,
            }
        ),
        encoding="utf-8",
    )

    parsed = load_run_config(config)

    assert parsed.provider.type == "thunderbird"
    assert parsed.provider.config == {"profile_path": "/some/path"}

    assert parsed.folders == [
        FolderNode(r"Inbox\Finance", ["finance"]),
        FolderNode(r"Inbox\HR", [])
    ]

    assert parsed.outputs.raw_root == tmp_path / "exports" / "raw"
    assert parsed.outputs.checkpoint_file == tmp_path / "exports" / "checkpoints" / "folder.json"
    assert parsed.outputs.markdown_root == tmp_path / "exports" / "markdown"
    assert parsed.outputs.save_raw is False

    assert parsed.extraction.cadence == "biweekly"
    assert parsed.extraction.start_date == "2020-01-01"
    assert parsed.extraction.end_date == "2020-01-31"
    assert parsed.extraction.manual is True
    assert parsed.extraction.max_windows == 2
    assert parsed.extraction.dry_run is True

    assert parsed.verbose is True


def test_load_run_config_rejects_invalid_max_windows(tmp_path: Path) -> None:
    config = tmp_path / "run.json"
    config.write_text(json.dumps({"extraction": {"max_windows": 0}}), encoding="utf-8")

    with pytest.raises(ValueError):
        load_run_config(config)


def test_load_run_config_accepts_utf8_bom_and_backwards_compat(tmp_path: Path) -> None:
    run_config = tmp_path / "run.json"
    payload = json.dumps({"folder": r"Shared Inbox\Product-A", "manual": True, "start_date": "2025-01-15"})
    run_config.write_bytes(payload.encode("utf-8-sig"))

    parsed = load_run_config(run_config)

    assert parsed.folders == [FolderNode(r"Shared Inbox\Product-A", [])]
    assert parsed.extraction.manual is True
    assert parsed.extraction.start_date == "2025-01-15"
