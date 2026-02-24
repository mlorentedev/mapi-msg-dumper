from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from mapi_msg_dumper.core.folders_config import load_folder_paths, normalize_folder_path


@dataclass(frozen=True)
class FileRunConfig:
    folder_paths: list[str]
    output_root: Path
    cadence: str
    start_date: str | None
    end_date: str | None
    manual: bool
    checkpoint_file: Path | None
    max_windows: int | None
    markdown_root: Path | None
    dry_run: bool
    verbose: bool


def load_run_config(config_path: Path) -> FileRunConfig:
    payload = json.loads(config_path.read_text(encoding="utf-8-sig"))
    if not isinstance(payload, dict):
        raise ValueError("Run config must be a JSON object.")

    folder_paths = _resolve_folder_paths(payload=payload, config_path=config_path)
    output_root = _read_path(payload=payload, config_path=config_path, key="output_root", default="exports")
    cadence = _read_optional_string(payload=payload, key="cadence") or "monthly"
    start_date = _read_optional_string(payload=payload, key="start_date")
    end_date = _read_optional_string(payload=payload, key="end_date")
    manual = _read_bool(payload=payload, key="manual", default=False)
    checkpoint_file = _read_optional_path(payload=payload, config_path=config_path, key="checkpoint_file")
    max_windows = _read_optional_positive_int(payload=payload, key="max_windows")
    markdown_root = _read_optional_path(payload=payload, config_path=config_path, key="markdown_root")
    dry_run = _read_bool(payload=payload, key="dry_run", default=False)
    verbose = _read_bool(payload=payload, key="verbose", default=False)

    return FileRunConfig(
        folder_paths=folder_paths,
        output_root=output_root,
        cadence=cadence,
        start_date=start_date,
        end_date=end_date,
        manual=manual,
        checkpoint_file=checkpoint_file,
        max_windows=max_windows,
        markdown_root=markdown_root,
        dry_run=dry_run,
        verbose=verbose,
    )


def _resolve_folder_paths(payload: dict[str, Any], config_path: Path) -> list[str]:
    if "folders" in payload:
        return load_folder_paths(config_path)

    folder = _read_optional_string(payload=payload, key="folder")
    return [normalize_folder_path(folder or "Inbox")]


def _read_bool(payload: dict[str, Any], key: str, default: bool) -> bool:
    value = payload.get(key, default)
    if isinstance(value, bool):
        return value
    raise ValueError(f"run config key '{key}' must be true or false.")


def _read_optional_string(payload: dict[str, Any], key: str) -> str | None:
    value = payload.get(key)
    if value is None:
        return None
    if not isinstance(value, str):
        raise ValueError(f"run config key '{key}' must be a string.")
    return value.strip() or None


def _read_path(payload: dict[str, Any], config_path: Path, key: str, default: str) -> Path:
    value = _read_optional_string(payload=payload, key=key) or default
    return _resolve_path(config_path=config_path, raw_path=value)


def _read_optional_path(payload: dict[str, Any], config_path: Path, key: str) -> Path | None:
    value = _read_optional_string(payload=payload, key=key)
    if value is None:
        return None
    return _resolve_path(config_path=config_path, raw_path=value)


def _read_optional_positive_int(payload: dict[str, Any], key: str) -> int | None:
    value = payload.get(key)
    if value is None:
        return None
    if not isinstance(value, int):
        raise ValueError(f"run config key '{key}' must be an integer.")
    if value < 1:
        raise ValueError(f"run config key '{key}' must be >= 1.")
    return value


def _resolve_path(config_path: Path, raw_path: str) -> Path:
    path = Path(raw_path)
    if path.is_absolute():
        return path
    return config_path.parent / path
