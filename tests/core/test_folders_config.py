import json
from pathlib import Path

from mapi_msg_dumper.core.folders_config import checkpoint_name_for_folder, load_folder_paths, normalize_folder_path


def test_normalize_folder_path_standardizes_separators() -> None:
    assert normalize_folder_path(r" Inbox/Team\\Archive ") == r"Inbox\Team\Archive"


def test_checkpoint_name_for_folder_is_stable() -> None:
    assert checkpoint_name_for_folder(r"Inbox\Team A\FY-2024") == "inbox-team-a-fy-2024"


def test_load_folder_paths_accepts_flat_list(tmp_path: Path) -> None:
    config = tmp_path / "folders.json"
    config.write_text(json.dumps({"folders": ["Inbox", r"Inbox\A", r"Inbox\A"]}), encoding="utf-8")

    assert load_folder_paths(config) == ["Inbox", r"Inbox\A"]


def test_load_folder_paths_expands_tree_relative_children(tmp_path: Path) -> None:
    config = tmp_path / "folders.json"
    config.write_text(
        json.dumps(
            {
                "folders": [
                    {
                        "path": "Inbox",
                        "include": False,
                        "children": [
                            {"path": "Clients", "children": ["2020", "2021"]},
                            "Finance",
                        ],
                    }
                ]
            }
        ),
        encoding="utf-8",
    )

    assert load_folder_paths(config) == [
        r"Inbox\Clients",
        r"Inbox\Clients\2020",
        r"Inbox\Clients\2021",
        r"Inbox\Finance",
    ]


def test_load_folder_paths_accepts_custom_root_paths(tmp_path: Path) -> None:
    config = tmp_path / "folders.json"
    config.write_text(json.dumps({"folders": [r"Shared Inbox\Product-A", r"Shared Inbox\Product-B"]}), encoding="utf-8")

    assert load_folder_paths(config) == [r"Shared Inbox\Product-A", r"Shared Inbox\Product-B"]


def test_load_folder_paths_accepts_utf8_bom(tmp_path: Path) -> None:
    config = tmp_path / "folders.json"
    payload = json.dumps({"folders": [r"Shared Inbox\Product-A"]})
    config.write_bytes(payload.encode("utf-8-sig"))

    assert load_folder_paths(config) == [r"Shared Inbox\Product-A"]
