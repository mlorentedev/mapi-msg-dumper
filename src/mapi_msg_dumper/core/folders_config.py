from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

_SAFE_TOKEN_CHARS = re.compile(r"[^a-z0-9]+")


@dataclass(frozen=True)
class FolderNode:
    path: str
    tags: list[str]


def load_folder_nodes(config_path: Path) -> list[FolderNode]:
    data = json.loads(config_path.read_text(encoding="utf-8-sig"))

    if isinstance(data, dict):
        folders_node = data.get("folders", [])
    else:
        folders_node = data

    if not isinstance(folders_node, list):
        raise ValueError("Folder config JSON must be an array or an object with a 'folders' array.")

    resolved: list[FolderNode] = []
    for node in folders_node:
        resolved.extend(_expand_node(node=node, parent=None, parent_tags=[]))

    if not resolved:
        raise ValueError("Folder config does not define any folder.")

    return _dedupe_keep_order(resolved)


def checkpoint_name_for_folder(folder_path: str) -> str:
    normalized = normalize_folder_path(folder_path)
    token = _SAFE_TOKEN_CHARS.sub("-", normalized.lower()).strip("-")
    return token or "folder"


def normalize_folder_path(folder_path: str) -> str:
    raw = folder_path.replace("/", "\\").strip().strip("\\")
    parts = [part.strip() for part in raw.split("\\") if part.strip()]
    if not parts:
        raise ValueError("Folder path cannot be empty.")
    return "\\".join(parts)


def _expand_node(node: Any, parent: str | None, parent_tags: list[str]) -> list[FolderNode]:
    if isinstance(node, str):
        return [FolderNode(path=_join_path(parent=parent, child_path=node), tags=list(parent_tags))]

    if not isinstance(node, dict):
        raise ValueError("Folder node must be a string path or an object with 'path' and optional 'children'/'tags'.")

    path_node = node.get("path")
    if not isinstance(path_node, str):
        raise ValueError("Folder node object must include a string 'path'.")

    include_current = node.get("include", True)
    if not isinstance(include_current, bool):
        raise ValueError("Folder node 'include' must be true or false.")

    current_tags = list(parent_tags)
    node_tags = node.get("tags", [])
    if not isinstance(node_tags, list):
        raise ValueError("Folder node 'tags' must be a list of strings.")
    for t in node_tags:
        if str(t) not in current_tags:
            current_tags.append(str(t))

    current = _join_path(parent=parent, child_path=path_node)
    children = node.get("children", [])
    if not isinstance(children, list):
        raise ValueError("Folder node 'children' must be an array.")

    resolved = [FolderNode(path=current, tags=current_tags)] if include_current else []
    for child in children:
        resolved.extend(_expand_node(node=child, parent=current, parent_tags=current_tags))
    return resolved


def _join_path(parent: str | None, child_path: str) -> str:
    child = normalize_folder_path(child_path)
    if parent is None:
        return child
    if _is_absolute_child_path(parent=parent, child=child):
        return child
    return normalize_folder_path(f"{parent}\\{child}")


def _is_absolute_child_path(parent: str, child: str) -> bool:
    if "\\" not in child:
        return False
    parent_root = normalize_folder_path(parent).split("\\", maxsplit=1)[0].lower()
    child_root = child.split("\\", maxsplit=1)[0].lower()
    return parent_root == child_root


def _dedupe_keep_order(nodes: list[FolderNode]) -> list[FolderNode]:
    seen: set[str] = set()
    ordered: list[FolderNode] = []
    for node in nodes:
        key = normalize_folder_path(node.path).lower()
        if key in seen:
            continue
        seen.add(key)
        ordered.append(FolderNode(path=normalize_folder_path(node.path), tags=node.tags))
    return ordered
