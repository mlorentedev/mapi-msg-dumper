from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any

from mapi_msg_dumper.core.folders_config import FolderNode
from mapi_msg_dumper.core.markdown import ExtractedEmail
from mapi_msg_dumper.core.planning import Window


class BaseExtractor(ABC):
    """Abstract strategy for extracting emails from a mail client or source."""

    @abstractmethod
    def connect(self) -> None:
        """Initialize connection to the data source."""
        pass

    @abstractmethod
    def get_messages(self, folder_node: FolderNode, window: Window) -> list[ExtractedEmail]:
        """Return a list of ExtractedEmail within the given time window for the given folder."""
        pass

    @abstractmethod
    def save_raw(self, email: ExtractedEmail, output_dir: Path) -> Path:
        """Save the raw format (.msg, .eml) to the output directory. Returns the path to the saved file."""
        pass
