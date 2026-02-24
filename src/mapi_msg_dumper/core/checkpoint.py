from __future__ import annotations

import json
from datetime import date
from pathlib import Path


def load_checkpoint(path: Path) -> date | None:
    if not path.exists():
        return None

    payload = json.loads(path.read_text(encoding="utf-8-sig"))
    next_start_date = payload.get("next_start_date")
    if next_start_date is None:
        return None
    return date.fromisoformat(next_start_date)


def save_checkpoint(path: Path, next_start_date: date) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {"next_start_date": next_start_date.isoformat()}
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
