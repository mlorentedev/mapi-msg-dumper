from __future__ import annotations

from datetime import date
from pathlib import Path

import typer
from rich.console import Console
from rich.table import Table

from mapi_msg_dumper.core.extractor import ExtractionSummary, run_extraction
from mapi_msg_dumper.core.folders_config import checkpoint_name_for_folder, normalize_folder_path
from mapi_msg_dumper.core.planning import normalize_cadence, parse_iso_date
from mapi_msg_dumper.core.run_config import load_run_config

app = typer.Typer(add_completion=False, no_args_is_help=True)
console = Console()


@app.command()
def extract(
    folder: str = typer.Option(
        "Inbox", help="Outlook folder path (examples: Inbox\\Subfolder, Shared Inbox\\Product-A)."
    ),
    run_config: Path | None = typer.Option(
        None, help="JSON run config. If set, values from this file drive execution."
    ),
    output_root: Path = typer.Option(Path("exports"), help="Root path for MSG output."),
    cadence: str = typer.Option("monthly", help="Auto batching cadence: monthly or biweekly."),
    start_date: str | None = typer.Option(None, help="Start date in YYYY-MM-DD."),
    end_date: str | None = typer.Option(None, help="End date in YYYY-MM-DD."),
    manual: bool = typer.Option(
        False, "--manual", help="Single-window mode. If not set, auto mode uses checkpoint resume."
    ),
    checkpoint_file: Path | None = typer.Option(
        None, help="Checkpoint file path. Defaults to <output-root>\\checkpoint.json."
    ),
    max_windows: int | None = typer.Option(
        None, min=1, help="Process at most N date windows per run (safe batching control)."
    ),
    markdown_root: Path | None = typer.Option(
        None, help="Optional root path to also write AI-friendly Markdown files."
    ),
    dry_run: bool = typer.Option(False, help="Evaluate and log without writing MSG files."),
    verbose: bool = typer.Option(False, "--verbose", help="Print detailed extraction progress."),
) -> None:
    try:
        if run_config is not None:
            config = load_run_config(run_config.resolve())
            parsed_start = _parse_optional_date(config.start_date)
            parsed_end = _parse_optional_date(config.end_date)
            normalized_cadence = normalize_cadence(config.cadence)
            folder_paths = config.folder_paths
            output_root = config.output_root
            manual = config.manual
            checkpoint_file = config.checkpoint_file
            max_windows = config.max_windows
            markdown_root = config.markdown_root
            dry_run = config.dry_run
            verbose = config.verbose
        else:
            parsed_start = _parse_optional_date(start_date)
            parsed_end = _parse_optional_date(end_date)
            normalized_cadence = normalize_cadence(cadence)
            folder_paths = [normalize_folder_path(folder)]

        folder_failures: list[tuple[str, str]] = []
        summary = ExtractionSummary()
        multi_folder = len(folder_paths) > 1

        for target_folder in folder_paths:
            if verbose and multi_folder:
                console.print(f"[cyan]Processing folder:[/cyan] {target_folder}")

            effective_checkpoint = _resolve_checkpoint_for_folder(
                checkpoint_file=checkpoint_file,
                output_root=output_root,
                folder_path=target_folder,
                multi_folder=multi_folder,
            )
            try:
                folder_summary = run_extraction(
                    folder_path=target_folder,
                    output_root=output_root,
                    cadence=normalized_cadence,
                    start_date=parsed_start,
                    end_date=parsed_end,
                    manual=manual,
                    checkpoint_path=effective_checkpoint,
                    dry_run=dry_run,
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
        output_root=output_root.resolve(),
        dry_run=dry_run,
        folders_requested=len(folder_paths),
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
    if not multi_folder:
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
    console.print(f"Output root: {output_root}")
    console.print(f"Success log: {output_root / 'logs' / 'success.csv'}")
    console.print(f"Error log:   {output_root / 'logs' / 'errors.csv'}")
