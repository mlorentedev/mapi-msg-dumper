from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from mapi_msg_dumper.core.folders_config import FolderNode, load_folder_nodes, normalize_folder_path


@dataclass(frozen=True)
class ProviderConfig:
    type: str
    config: dict[str, Any]


@dataclass(frozen=True)
class ExtractionConfig:
    start_date: str | None
    end_date: str | None
    cadence: str
    max_windows: int | None
    dry_run: bool
    manual: bool


@dataclass(frozen=True)
class OutputsConfig:
    markdown_root: Path | None
    raw_root: Path | None
    save_raw: bool
    checkpoint_file: Path | None


@dataclass(frozen=True)
class FileRunConfig:
    provider: ProviderConfig
    extraction: ExtractionConfig
    outputs: OutputsConfig
    folders: list[FolderNode]
    verbose: bool


def load_run_config(config_path: Path) -> FileRunConfig:
    payload = json.loads(config_path.read_text(encoding="utf-8-sig"))
    if not isinstance(payload, dict):
        raise ValueError("Run config must be a JSON object.")

    # 1. Provider
    provider_node = payload.get("provider", {})
    if not isinstance(provider_node, dict):
        raise ValueError("run config 'provider' must be an object.")
    provider_type = _read_optional_string(provider_node, "type") or "outlook"
    provider_config = provider_node.get("config", {})
    if not isinstance(provider_config, dict):
        raise ValueError("run config 'provider.config' must be an object.")

    # 2. Extraction
    ext_node = payload.get("extraction", payload)  # fallback to root for backwards compat during transition
    if not isinstance(ext_node, dict):
        raise ValueError("run config 'extraction' must be an object.")

    start_date = _read_optional_string(ext_node, "start_date")
    end_date = _read_optional_string(ext_node, "end_date")
    cadence = _read_optional_string(ext_node, "cadence") or "monthly"
    max_windows = _read_optional_positive_int(ext_node, "max_windows")
    dry_run = _read_bool(ext_node, "dry_run", default=False)
    manual = _read_bool(ext_node, "manual", default=False)

    # 3. Outputs
    out_node = payload.get("outputs", payload)
    if not isinstance(out_node, dict):
        raise ValueError("run config 'outputs' must be an object.")

    md_root = _read_optional_path(out_node, config_path, "markdown_root")
    raw_root = _read_optional_path(out_node, config_path, "raw_root")

    # Backwards compatibility: output_root mapped to raw_root if raw_root is missing
    if raw_root is None:
        raw_root = _read_optional_path(out_node, config_path, "output_root")
        if raw_root is None:
            raw_root = _resolve_path(config_path, "exports")

    save_raw = _read_bool(out_node, "save_raw", default=True)
    checkpoint_file = _read_optional_path(out_node, config_path, "checkpoint_file")

    # 4. Folders
    folders = _resolve_folder_nodes(payload, config_path)

    # 5. Root flags
    verbose = _read_bool(payload, "verbose", default=False)

    return FileRunConfig(
        provider=ProviderConfig(type=provider_type, config=provider_config),
        extraction=ExtractionConfig(
            start_date=start_date,
            end_date=end_date,
            cadence=cadence,
            max_windows=max_windows,
            dry_run=dry_run,
            manual=manual,
        ),
        outputs=OutputsConfig(
            markdown_root=md_root,
            raw_root=raw_root,
            save_raw=save_raw,
            checkpoint_file=checkpoint_file,
        ),
        folders=folders,
        verbose=verbose,
    )


def _resolve_folder_nodes(payload: dict[str, Any], config_path: Path) -> list[FolderNode]:
    if "folders" in payload:
        return load_folder_nodes(config_path)

    folder = _read_optional_string(payload, "folder")
    return [FolderNode(path=normalize_folder_path(folder or "Inbox"), tags=[])]


def _read_bool(payload: dict[str, Any], key: str, default: bool) -> bool:
    value = payload.get(key, default)
    if isinstance(value, bool):
        return value
    raise ValueError(f"config key '{key}' must be true or false.")


def _read_optional_string(payload: dict[str, Any], key: str) -> str | None:
    value = payload.get(key)
    if value is None:
        return None
    if not isinstance(value, str):
        raise ValueError(f"config key '{key}' must be a string.")
    return value.strip() or None


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
        raise ValueError(f"config key '{key}' must be an integer.")
    if value < 1:
        raise ValueError(f"config key '{key}' must be >= 1.")
    return value


def _resolve_path(config_path: Path, raw_path: str) -> Path:
    path = Path(raw_path)
    if path.is_absolute():
        return path
    return config_path.parent / path
