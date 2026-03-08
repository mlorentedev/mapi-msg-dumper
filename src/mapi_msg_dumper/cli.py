from __future__ import annotations

from datetime import date
from pathlib import Path

import typer
from rich.console import Console
from rich.table import Table

from mapi_msg_dumper.core.extractor import ExtractionSummary, run_extraction
from mapi_msg_dumper.core.extractors.base import BaseExtractor
from mapi_msg_dumper.core.extractors.outlook import OutlookExtractor
from mapi_msg_dumper.core.extractors.thunderbird import ThunderbirdExtractor
from mapi_msg_dumper.core.folders_config import FolderNode, checkpoint_name_for_folder, normalize_folder_path
from mapi_msg_dumper.core.planning import normalize_cadence, parse_iso_date
from mapi_msg_dumper.core.run_config import load_run_config

app = typer.Typer(add_completion=False, no_args_is_help=True)
console = Console()


def _create_extractor(provider_type: str, provider_config: dict) -> BaseExtractor:
    if provider_type == "outlook":
        return OutlookExtractor()
    if provider_type == "thunderbird":
        profile = provider_config.get("profile_path")
        if not profile:
            raise ValueError("Thunderbird provider requires 'profile_path' in config.")
        return ThunderbirdExtractor(profile_path=profile)
    raise ValueError(f"Unknown provider type: {provider_type}")


@app.command()
def extract(
    folder: str = typer.Option(
        "Inbox", help="Outlook folder path (examples: Inbox\\Subfolder, Shared Inbox\\Product-A)."
    ),
    run_config: Path | None = typer.Option(
        None, help="JSON run config. If set, values from this file drive execution."
    ),
    output_root: Path = typer.Option(Path("exports"), help="Root path for raw output (MSG/EML)."),
    cadence: str = typer.Option("monthly", help="Auto batching cadence: monthly or biweekly."),
    start_date: str | None = typer.Option(None, help="Start date in YYYY-MM-DD."),
    end_date: str | None = typer.Option(None, help="End date in YYYY-MM-DD."),
    manual: bool = typer.Option(
        False, "--manual", help="Single-window mode. If not set, auto mode uses checkpoint resume."
    ),
    checkpoint_file: Path | None = typer.Option(
        None, help="Checkpoint file path. Defaults to <raw-root>\\checkpoints\\<folder>.json."
    ),
    max_windows: int | None = typer.Option(
        None, min=1, help="Process at most N date windows per run (safe batching control)."
    ),
    markdown_root: Path | None = typer.Option(
        None, help="Optional root path to also write AI-friendly Markdown files."
    ),
    dry_run: bool = typer.Option(False, help="Evaluate and log without writing files."),
    verbose: bool = typer.Option(False, "--verbose", help="Print detailed extraction progress."),
) -> None:
    try:
        if run_config is not None:
            config = load_run_config(run_config.resolve())
            parsed_start = _parse_optional_date(config.extraction.start_date)
            parsed_end = _parse_optional_date(config.extraction.end_date)
            normalized_cadence = normalize_cadence(config.extraction.cadence)
            folder_nodes = config.folders

            raw_root = config.outputs.raw_root or output_root
            manual = config.extraction.manual
            checkpoint_file = config.outputs.checkpoint_file
            max_windows = config.extraction.max_windows
            markdown_root = config.outputs.markdown_root
            dry_run = config.extraction.dry_run
            save_raw = config.outputs.save_raw
            verbose = config.verbose
            extractor = _create_extractor(config.provider.type, config.provider.config)
        else:
            parsed_start = _parse_optional_date(start_date)
            parsed_end = _parse_optional_date(end_date)
            normalized_cadence = normalize_cadence(cadence)
            folder_nodes = [FolderNode(path=normalize_folder_path(folder), tags=[])]
            raw_root = output_root
            save_raw = True
            extractor = OutlookExtractor()

        folder_failures: list[tuple[str, str]] = []
        summary = ExtractionSummary()
        multi_folder = len(folder_nodes) > 1

        for target_node in folder_nodes:
            target_folder = target_node.path
            if verbose and multi_folder:
                console.print(f"[cyan]Processing folder:[/cyan] {target_folder}")

            effective_checkpoint = _resolve_checkpoint_for_folder(
                checkpoint_file=checkpoint_file,
                output_root=raw_root,
                folder_path=target_folder,
                multi_folder=multi_folder,
            )
            try:
                folder_summary = run_extraction(
                    extractor=extractor,
                    folder_node=target_node,
                    output_root=raw_root,
                    cadence=normalized_cadence,
                    start_date=parsed_start,
                    end_date=parsed_end,
                    manual=manual,
                    checkpoint_path=effective_checkpoint,
                    dry_run=dry_run,
                    save_raw=save_raw,
                    markdown_root=markdown_root,
                    verbose=verbose,
                    max_windows=max_windows,
                )
                summary.merge(folder_summary)
            except Exception as exc:
                folder_failures.append((target_folder, str(exc)))
                console.print(f"[red]Folder failed:[/red] {target_folder} -> {exc}")

    except Exception as exc:
        console.print(f"[red]Extraction failed:[/red] {exc}")
        raise typer.Exit(code=1) from exc

    _print_summary(
        summary=summary,
        output_root=raw_root.resolve(),
        dry_run=dry_run,
        folders_requested=len(folder_nodes),
        folders_failed=len(folder_failures),
    )
    if folder_failures:
        raise typer.Exit(code=1)


def _parse_optional_date(value: str | None) -> date | None:
    if value is None:
        return None
    return parse_iso_date(value)


def _resolve_checkpoint_for_folder(
    checkpoint_file: Path | None, output_root: Path, folder_path: str, multi_folder: bool
) -> Path | None:
    if not multi_folder and checkpoint_file is not None:
        return checkpoint_file

    token = checkpoint_name_for_folder(folder_path)
    if checkpoint_file is None:
        return output_root / "checkpoints" / f"{token}.json"
    if checkpoint_file.suffix.lower() == ".json":
        return checkpoint_file.with_name(f"{checkpoint_file.stem}.{token}{checkpoint_file.suffix}")
    return checkpoint_file / f"{token}.json"


def _print_summary(
    summary: ExtractionSummary, output_root: Path, dry_run: bool, folders_requested: int, folders_failed: int
) -> None:
    table = Table(title="Extraction Summary")
    table.add_column("Metric")
    table.add_column("Value", justify="right")

    table.add_row("Folders requested", str(folders_requested))
    table.add_row("Folders failed", str(folders_failed))
    table.add_row("Windows processed", str(summary.windows_processed))
    table.add_row("Exported", str(summary.exported))
    table.add_row("Markdown written", str(summary.markdown_written))
    table.add_row("Skipped (existing)", str(summary.skipped_existing))
    table.add_row("Skipped (non-mail)", str(summary.skipped_non_mail))
    table.add_row("Failed", str(summary.failed))
    table.add_row("Dry run", str(dry_run).lower())

    console.print(table)
    console.print(f"Raw Output root: {output_root}")
    console.print(f"Success log: {output_root / 'logs' / 'success.csv'}")
    console.print(f"Error log:   {output_root / 'logs' / 'errors.csv'}")
