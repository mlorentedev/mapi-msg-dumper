from __future__ import annotations

from pathlib import Path
from typing import Any

from mapi_msg_dumper.core.extractors.base import BaseExtractor
from mapi_msg_dumper.core.folders_config import FolderNode
from mapi_msg_dumper.core.markdown import ExtractedEmail
from mapi_msg_dumper.core.planning import Window


class ThunderbirdExtractor(BaseExtractor):
    def __init__(self, profile_path: str) -> None:
        self.profile_path = Path(profile_path)

    def connect(self) -> None:
        if not self.profile_path.exists():
            raise ValueError(f"Thunderbird profile path does not exist: {self.profile_path}")

    def get_messages(self, folder_node: FolderNode, window: Window) -> list[ExtractedEmail]:
        # TODO: Implement local .mbox parsing
        return []

    def save_raw(self, email: ExtractedEmail, output_dir: Any) -> Any:
        # TODO: Implement writing .eml
        pass
