from __future__ import annotations

import csv
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from mapi_msg_dumper.core.checkpoint import load_checkpoint, save_checkpoint
from mapi_msg_dumper.core.extractors.base import BaseExtractor
from mapi_msg_dumper.core.filenames import markdown_file_path
from mapi_msg_dumper.core.folders_config import FolderNode
from mapi_msg_dumper.core.markdown import render_email_markdown
from mapi_msg_dumper.core.planning import (
    Cadence,
    Window,
    apply_window_limit,
    build_auto_windows,
    build_manual_window,
)


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
    extractor: BaseExtractor,
    folder_node: FolderNode,
    output_root: Path,
    cadence: Cadence,
    start_date: date | None,
    end_date: date | None,
    manual: bool,
    checkpoint_path: Path | None,
    dry_run: bool,
    save_raw: bool = True,
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
            f"[mapi-msg-dumper] mode={mode} folder={folder_node.path} windows={len(windows)} "
            f"window_limit={max_windows if max_windows is not None else 'none'} dry_run={str(dry_run).lower()}"
        )

    extractor.connect()

    success_log = destination / "logs" / "success.csv"
    error_log = destination / "logs" / "errors.csv"

    for window in windows:
        if verbose:
            print(f"[mapi-msg-dumper] window {window.start.isoformat()} -> {window.end.isoformat()}")
        window_summary = _export_window(
            extractor,
            folder_node,
            destination,
            window,
            success_log,
            error_log,
            dry_run,
            save_raw,
            verbose,
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


def _export_window(
    extractor: BaseExtractor,
    folder_node: FolderNode,
    output_root: Path,
    window: Window,
    success_log: Path,
    error_log: Path,
    dry_run: bool,
    save_raw: bool,
    verbose: bool,
    markdown_root: Path | None,
) -> ExtractionSummary:
    summary = ExtractionSummary()
    try:
        messages = extractor.get_messages(folder_node, window)
    except Exception as exc:
        if verbose:
            print(f"[mapi-msg-dumper] error extracting messages from window {window.start.isoformat()}: {exc}")
        return summary

    for email in messages:
        try:
            raw_path = None
            if save_raw:
                if dry_run:
                    if verbose:
                        print(f"[mapi-msg-dumper] simulated save raw entry_id={email.entry_id}")
                    summary.exported += 1
                else:
                    raw_path = extractor.save_raw(email, output_root)
                    summary.exported += 1
                    if verbose:
                        print(f"[mapi-msg-dumper] saved raw {raw_path}")

                    _append_csv(
                        success_log,
                        ["window_start", "window_end", "entry_id", "saved_path", "dry_run"],
                        {
                            "window_start": window.start.isoformat(),
                            "window_end": window.end.isoformat(),
                            "entry_id": email.entry_id,
                            "saved_path": str(raw_path),
                            "dry_run": str(dry_run).lower(),
                        },
                    )

            if markdown_root is not None:
                md_path = markdown_file_path(markdown_root, email.received_at, email.subject, email.entry_id)
                if md_path.exists():
                    if verbose:
                        print(f"[mapi-msg-dumper] skip existing markdown {md_path}")
                elif dry_run:
                    if verbose:
                        print(f"[mapi-msg-dumper] simulated markdown {md_path}")
                else:
                    markdown = render_email_markdown(email, source_raw_path=raw_path)
                    md_path.parent.mkdir(parents=True, exist_ok=True)
                    md_path.write_text(markdown, encoding="utf-8")
                    summary.markdown_written += 1
                    if verbose:
                        print(f"[mapi-msg-dumper] saved markdown {md_path}")

        except Exception as exc:
            summary.failed += 1
            if verbose:
                print(f"[mapi-msg-dumper] error entry_id={email.entry_id} subject={email.subject!r}: {exc}")
            _append_csv(
                error_log,
                ["window_start", "window_end", "entry_id", "subject", "error"],
                {
                    "window_start": window.start.isoformat(),
                    "window_end": window.end.isoformat(),
                    "entry_id": email.entry_id,
                    "subject": email.subject,
                    "error": str(exc),
                },
            )

    return summary


def _append_csv(path: Path, fieldnames: list[str], row: dict[str, str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    file_exists = path.exists()
    with path.open("a", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        if not file_exists:
            writer.writeheader()
        writer.writerow(row)
