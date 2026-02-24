# mapi-msg-dumper

Local Outlook extraction engine using COM (`pywin32`) to export historical emails as `.msg` and optional AI-friendly Markdown.

## Why

When Microsoft Graph is blocked by tenant policy, this tool works directly against the authenticated Outlook desktop session.

## Install

```powershell
poetry install
```

## Single-command operation

Use one config file only (`run.json`), then always run:

```powershell
poetry run mapi-msg-dumper --run-config .\run.json
```

`run.json` is local-only and ignored by git (`.gitignore`).
You can bootstrap it from the sanitized template:

```powershell
Copy-Item .\run.example.json .\run.json
```

## Recommended workflow

1. **Edit only `run.json`** (folders + date range + output settings).
2. Set `"dry_run": true` for a safe validation pass.
3. Run the command above.
4. If folder validation looks good, set `"dry_run": false` and run again.
5. Move to next month by changing only `start_date` and `end_date`.

## `run.example.json` reference

```json
{
  "folders": [
    "Shared Inbox\\Product-A",
    "Shared Inbox\\Product-B",
    "Inbox\\Project-X"
  ],
  "cadence": "monthly",
  "manual": true,
  "start_date": "2025-01-01",
  "end_date": "2025-01-31",
  "output_root": "./exports",
  "checkpoint_file": "./exports/checkpoints/history.json",
  "markdown_root": "./exports/markdown",
  "max_windows": 1,
  "dry_run": false,
  "verbose": true
}
```

### Supported keys

- `folder` or `folders`
- `output_root`
- `cadence` (`monthly` or `biweekly`)
- `start_date`, `end_date` (`YYYY-MM-DD`)
- `manual`
- `checkpoint_file`
- `max_windows`
- `markdown_root`
- `dry_run`
- `verbose`

Notes:
- `folders` supports flat paths and tree nodes (`path` + optional `children`).
- Folder roots may be `Inbox\...` or other top-level mailbox folders (for example `Shared Inbox\Product-A`).
- UTF-8 BOM JSON is supported (PowerShell `Set-Content` compatible).

## Output layout (single export root)

All runtime artifacts can live under one root (`exports`):

```text
exports/
├── YYYY/MM/*.msg
├── markdown/YYYY/MM/*.md
├── logs/success.csv
├── logs/errors.csv
└── checkpoints/*.json
```

## Quick checks after each run

```powershell
Import-Csv .\exports\logs\success.csv | Select-Object -Last 10
Import-Csv .\exports\logs\errors.csv | Select-Object -Last 10
Get-ChildItem .\exports -Recurse -Filter *.msg | Select-Object -First 10 FullName
```

## Common issues

- `Folder failed ... segment ... not found`: the folder path does not exist in Outlook exactly as written.
- `Dry run = true` with no files: expected behavior (simulation mode).
- Missing old emails: Outlook may still be syncing cached history; wait for sync completion.

## Release strategy (GitHub)

This project uses:
- Conventional Commits (`feat:`, `fix:`, `docs:`, `refactor:`, `test:`, `chore:`)
- CI workflow (`.github/workflows/ci.yml`) for `ruff`, `mypy`, and `pytest`
- `release-please` (`release-please-config.json` + `.release-please-manifest.json`) for semantic versioning and changelog PRs
- Release workflow (`.github/workflows/release.yml`) to build Python distributions and attach them to GitHub Releases
