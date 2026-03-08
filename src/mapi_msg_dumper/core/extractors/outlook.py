from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

from mapi_msg_dumper.core.extractors.base import BaseExtractor
from mapi_msg_dumper.core.filenames import message_file_path
from mapi_msg_dumper.core.folders_config import FolderNode
from mapi_msg_dumper.core.markdown import ExtractedEmail
from mapi_msg_dumper.core.planning import Window, build_received_filter

OL_FOLDER_INBOX = 6
OL_MAIL_ITEM = 43
OL_MSG_UNICODE = 3


class OutlookExtractor(BaseExtractor):
    def __init__(self) -> None:
        self.namespace: Any = None

    def connect(self) -> None:
        import win32com.client  # type: ignore[import-untyped]

        self.namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    def get_messages(self, folder_node: FolderNode, window: Window) -> list[ExtractedEmail]:
        if not self.namespace:
            raise RuntimeError("Must call connect() before extracting.")

        folder = self._resolve_folder(folder_node.path)
        items = folder.Items
        items.Sort("[ReceivedTime]", False)
        scoped = items.Restrict(build_received_filter(window))

        extracted = []
        item = scoped.GetFirst()
        while item is not None:
            try:
                if int(getattr(item, "Class", 0)) != OL_MAIL_ITEM:
                    continue

                entry_id = str(getattr(item, "EntryID", ""))
                subject = str(getattr(item, "Subject", ""))
                received_at = self._received_datetime(item)
                body = self._safe_text(getattr(item, "Body", ""))

                email = ExtractedEmail(
                    received_at=received_at,
                    subject=subject,
                    sender_name=self._safe_text(getattr(item, "SenderName", "")),
                    sender_email=self._safe_text(getattr(item, "SenderEmailAddress", "")),
                    to=self._safe_text(getattr(item, "To", "")),
                    cc=self._safe_text(getattr(item, "CC", "")),
                    entry_id=entry_id,
                    folder_path=folder_node.path,
                    body_text=body,
                    tags=folder_node.tags,
                )
                # Keep a reference to the COM object for saving later
                email._com_item = item  # type: ignore[attr-defined]
                extracted.append(email)
            except Exception as exc:
                # We skip failing items here, or we can raise to abort window
                # Let's bubble up or just wrap in a specialized error
                raise RuntimeError(f"Failed extracting item from COM: {exc}") from exc
            finally:
                item = scoped.GetNext()

        return extracted

    def save_raw(self, email: ExtractedEmail, output_dir: Path) -> Path:
        msg_path = message_file_path(output_dir, email.received_at, email.subject, email.entry_id)
        if msg_path.exists():
            return msg_path

        com_item = getattr(email, "_com_item", None)
        if com_item is None:
            raise ValueError("OutlookExtractor requires _com_item on email to save raw.")

        msg_path.parent.mkdir(parents=True, exist_ok=True)
        com_item.SaveAs(str(msg_path), OL_MSG_UNICODE)
        return msg_path

    def _resolve_folder(self, folder_path: str) -> Any:
        parts = [part.strip() for part in folder_path.split("\\") if part.strip()]
        if not parts:
            raise ValueError("Folder path cannot be empty.")

        if parts[0].lower() == "inbox":
            return self._walk_folder_path(self.namespace.GetDefaultFolder(OL_FOLDER_INBOX), parts[1:])

        try:
            root_folder = self.namespace.DefaultStore.GetRootFolder()
        except Exception as exc:
            raise ValueError("Cannot resolve default Outlook store root folder.") from exc
        return self._walk_folder_path(root_folder, parts)

    def _walk_folder_path(self, start_folder: Any, segments: list[str]) -> Any:
        folder = start_folder
        for segment in segments:
            folder = self._get_child_folder(folder, segment)
        return folder

    def _get_child_folder(self, parent: Any, segment: str) -> Any:
        try:
            return parent.Folders.Item(segment)
        except Exception:
            count = int(getattr(parent.Folders, "Count", 0))
            for index in range(1, count + 1):
                child = parent.Folders.Item(index)
                if str(getattr(child, "Name", "")).strip().lower() == segment.lower():
                    return child
        parent_name = str(getattr(parent, "Name", "<root>"))
        raise ValueError(f"Folder segment '{segment}' not found under '{parent_name}'.")

    def _received_datetime(self, item: Any) -> datetime:
        received = getattr(item, "ReceivedTime", None)
        if isinstance(received, datetime):
            return received.replace(tzinfo=None)
        raise ValueError("Item has no valid ReceivedTime.")

    def _safe_text(self, value: Any) -> str:
        if value is None:
            return ""
        return str(value)
