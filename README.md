# mapi-msg-dumper

Local email extraction engine (Outlook/Thunderbird) designed for batch processing historical data and converting it into AI-friendly Markdown.

## Why

When Microsoft Graph is blocked by tenant policy or you need to process local archives (Thunderbird MBOX) without configuring IMAP/OAuth, this tool works directly against the local data stores.

## Features

- **Strategy Pattern Extractors**: Native support for Outlook (via MAPI/COM) and Thunderbird (local profiles).
- **Batch Resiliency**: Time-window planning (monthly/biweekly) and checkpointing to resume long extractions.
- **AI-Optimized**: Generates clean Markdown with YAML frontmatter and folder-based tagging.
- **Offline-First**: No credentials required; if your local client can read it, this tool can extract it.

## Install

```bash
poetry install
```

## Operation

The tool is driven by a `run.json` configuration file:

```bash
poetry run mapi-msg-dumper --run-config .\run.json
```

### `run.json` reference

```json
{
  "provider": {
    "type": "outlook",
    "config": {}
  },
  "extraction": {
    "start_date": "2025-01-01",
    "end_date": "2025-01-31",
    "cadence": "monthly",
    "manual": true,
    "max_windows": 1,
    "dry_run": false
  },
  "outputs": {
    "raw_root": "./exports/raw",
    "markdown_root": "./exports/markdown",
    "save_raw": true
  },
  "folders": [
    {
      "path": "Inbox\\Project-X",
      "tags": ["project-x", "internal"]
    },
    "Shared Inbox\\Product-A"
  ],
  "verbose": true
}
```

### Configuration Sections

- **`provider`**: 
    - `type`: `"outlook"` or `"thunderbird"`.
    - `config`: Provider-specific settings (e.g., `profile_path` for Thunderbird).
- **`extraction`**: Controls the date range and batching logic.
- **`outputs`**:
    - `raw_root`: Where to save original `.msg` or `.eml` files.
    - `markdown_root`: Where to save the generated `.md` files.
    - `save_raw`: Set to `false` if you only need Markdown.
- **`folders`**: List of folder paths to process. Supports flat strings or objects with `tags` and `children` for tree-based configuration.

## Release Strategy

This project uses:
- **Conventional Commits**: `feat:`, `fix:`, `docs:`, `refactor:`, `test:`, `chore:`.
- **Release-Please**: Automated versioning (vX.X.X) and changelog generation.
- **CI/CD**: `ruff`, `mypy`, and `pytest` validation on every PR.
