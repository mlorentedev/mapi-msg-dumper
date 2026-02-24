from __future__ import annotations

import csv
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

import win32com.client  # type: ignore[import-untyped]

from mapi_msg_dumper.core.checkpoint import load_checkpoint, save_checkpoint
from mapi_msg_dumper.core.filenames import markdown_file_path, message_file_path
from mapi_msg_dumper.core.markdown import MarkdownEmail, render_email_markdown
from mapi_msg_dumper.core.planning import (
    Cadence,
    Window,
    apply_window_limit,
    build_auto_windows,
    build_manual_window,
    build_received_filter,
)

OL_FOLDER_INBOX = 6
OL_MAIL_ITEM = 43
OL_MSG_UNICODE = 3


@dataclass
class ExtractionSummary:
    windows_processed: int = 0
    exported: int = 0
    markdown_written: int = 0
    skipped_existing: int = 0
    skipped_non_mail: int = 0
    failed: int = 0

    def merge(self, other: "ExtractionSummary") -> None:
        self.windows_processed += other.windows_processed
        self.exported += other.exported
        self.markdown_written += other.markdown_written
        self.skipped_existing += other.skipped_existing
        self.skipped_non_mail += other.skipped_non_mail
        self.failed += other.failed


def run_extraction(
    folder_path: str,
    output_root: Path,
    cadence: Cadence,
    start_date: date | None,
    end_date: date | None,
    manual: bool,
    checkpoint_path: Path | None,
    dry_run: bool,
    markdown_root: Path | None = None,
    verbose: bool = False,
    max_windows: int | None = None,
) -> ExtractionSummary:
    if manual and end_date is None:
        raise ValueError("Manual mode requires --end-date.")

    destination = output_root.resolve()
    effective_end = end_date or datetime.now().date()
    checkpoint = checkpoint_path or destination / "checkpoint.json"

    windows = _build_windows(cadence, start_date, effective_end, manual, checkpoint)
    windows = apply_window_limit(windows, max_windows)
    summary = ExtractionSummary()
    if not windows:
        return summary

    if verbose:
        mode = "manual" if manual else f"auto/{cadence}"
        print(
            f"[mapi-msg-dumper] mode={mode} folder={folder_path} windows={len(windows)} "
            f"window_limit={max_windows if max_windows is not None else 'none'} dry_run={str(dry_run).lower()}"
        )

    namespace = _connect_namespace()
    folder = _resolve_folder(namespace, folder_path)

    success_log = destination / "logs" / "success.csv"
    error_log = destination / "logs" / "errors.csv"

    for window in windows:
        if verbose:
            print(f"[mapi-msg-dumper] window {window.start.isoformat()} -> {window.end.isoformat()}")
        window_summary = _export_window(
            folder,
            destination,
            window,
            success_log,
            error_log,
            dry_run,
            verbose,
            folder_path,
            markdown_root.resolve() if markdown_root is not None else None,
        )
        window_summary.windows_processed = 1
        summary.merge(window_summary)
        if not manual and not dry_run:
            save_checkpoint(checkpoint, window.end.date())
            if verbose:
                print(f"[mapi-msg-dumper] checkpoint updated to {window.end.date().isoformat()}")

    return summary


def _build_windows(
    cadence: Cadence,
    start_date: date | None,
    end_date: date,
    manual: bool,
    checkpoint_path: Path,
) -> list[Window]:
    if manual:
        if start_date is None:
            raise ValueError("Manual mode requires --start-date.")
        return [build_manual_window(start_date, end_date)]

    resume_from = load_checkpoint(checkpoint_path)
    first_start = resume_from or start_date
    if first_start is None:
        raise ValueError("Auto mode needs --start-date on first run when no checkpoint file exists.")
    return build_auto_windows(first_start, end_date, cadence)


def _connect_namespace() -> Any:
    return win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")


def _resolve_folder(namespace: Any, folder_path: str) -> Any:
    parts = [part.strip() for part in folder_path.split("\\") if part.strip()]
    if not parts:
        raise ValueError("Folder path cannot be empty.")

    if parts[0].lower() == "inbox":
        return _walk_folder_path(namespace.GetDefaultFolder(OL_FOLDER_INBOX), parts[1:])

    try:
        root_folder = namespace.DefaultStore.GetRootFolder()
    except Exception as exc:
        raise ValueError("Cannot resolve default Outlook store root folder.") from exc
    return _walk_folder_path(root_folder, parts)


def _walk_folder_path(start_folder: Any, segments: list[str]) -> Any:
    folder = start_folder
    for segment in segments:
        folder = _get_child_folder(folder, segment)
    return folder


def _get_child_folder(parent: Any, segment: str) -> Any:
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


def _export_window(
    folder: Any,
    output_root: Path,
    window: Window,
    success_log: Path,
    error_log: Path,
    dry_run: bool,
    verbose: bool,
    folder_path: str,
    markdown_root: Path | None,
) -> ExtractionSummary:
    summary = ExtractionSummary()

    items = folder.Items
    items.Sort("[ReceivedTime]", False)
    scoped = items.Restrict(build_received_filter(window))

    item = scoped.GetFirst()
    while item is not None:
        entry_id = ""
        subject = ""
        try:
            if int(getattr(item, "Class", 0)) != OL_MAIL_ITEM:
                summary.skipped_non_mail += 1
                continue

            entry_id = str(getattr(item, "EntryID", ""))
            subject = str(getattr(item, "Subject", ""))
            received_at = _received_datetime(item)
            msg_path = message_file_path(output_root, received_at, subject, entry_id)

            if msg_path.exists():
                summary.skipped_existing += 1
                if verbose:
                    print(f"[mapi-msg-dumper] skip existing {msg_path}")
            else:
                if not dry_run:
                    msg_path.parent.mkdir(parents=True, exist_ok=True)
                    item.SaveAs(str(msg_path), OL_MSG_UNICODE)

                summary.exported += 1
                if verbose:
                    action = "simulated save" if dry_run else "saved"
                    print(f"[mapi-msg-dumper] {action} {msg_path}")
                _append_csv(
                    success_log,
                    ["window_start", "window_end", "entry_id", "saved_path", "dry_run"],
                    {
                        "window_start": window.start.isoformat(),
                        "window_end": window.end.isoformat(),
                        "entry_id": entry_id,
                        "saved_path": str(msg_path),
                        "dry_run": str(dry_run).lower(),
                    },
                )

            if markdown_root is not None:
                md_path = markdown_file_path(markdown_root, received_at, subject, entry_id)
                if md_path.exists():
                    if verbose:
                        print(f"[mapi-msg-dumper] skip existing markdown {md_path}")
                elif dry_run:
                    if verbose:
                        print(f"[mapi-msg-dumper] simulated markdown {md_path}")
                else:
                    body = _safe_text(getattr(item, "Body", ""))
                    markdown = render_email_markdown(
                        MarkdownEmail(
                            received_at=received_at,
                            subject=subject,
                            sender_name=_safe_text(getattr(item, "SenderName", "")),
                            sender_email=_safe_text(getattr(item, "SenderEmailAddress", "")),
                            to=_safe_text(getattr(item, "To", "")),
                            cc=_safe_text(getattr(item, "CC", "")),
                            entry_id=entry_id,
                            source_msg_path=msg_path,
                            folder_path=folder_path,
                        ),
                        body=body,
                    )
                    md_path.parent.mkdir(parents=True, exist_ok=True)
                    md_path.write_text(markdown, encoding="utf-8")
                    summary.markdown_written += 1
                    if verbose:
                        print(f"[mapi-msg-dumper] saved markdown {md_path}")
        except Exception as exc:
            summary.failed += 1
            if verbose:
                print(f"[mapi-msg-dumper] error entry_id={entry_id or 'unknown'} subject={subject!r}: {exc}")
            _append_csv(
                error_log,
                ["window_start", "window_end", "entry_id", "subject", "error"],
                {
                    "window_start": window.start.isoformat(),
                    "window_end": window.end.isoformat(),
                    "entry_id": entry_id,
                    "subject": subject,
                    "error": str(exc),
                },
            )
        finally:
            item = scoped.GetNext()

    return summary


def _received_datetime(item: Any) -> datetime:
    received = getattr(item, "ReceivedTime", None)
    if isinstance(received, datetime):
        return received.replace(tzinfo=None)
    raise ValueError("Item has no valid ReceivedTime.")


def _safe_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value)


def _append_csv(path: Path, fieldnames: list[str], row: dict[str, str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    file_exists = path.exists()
    with path.open("a", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        if not file_exists:
            writer.writeheader()
        writer.writerow(row)
