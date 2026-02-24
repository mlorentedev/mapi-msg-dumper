from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from typing import Literal, cast

Cadence = Literal["monthly", "biweekly"]


@dataclass(frozen=True)
class Window:
    start: datetime
    end: datetime  # exclusive


def parse_iso_date(value: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise ValueError(f"Invalid date '{value}'. Use YYYY-MM-DD.") from exc


def normalize_cadence(value: str) -> Cadence:
    normalized = value.strip().lower()
    if normalized not in {"monthly", "biweekly"}:
        raise ValueError("Cadence must be 'monthly' or 'biweekly'.")
    return cast(Cadence, normalized)


def build_auto_windows(start_day: date, end_day: date, cadence: Cadence) -> list[Window]:
    if start_day > end_day:
        raise ValueError("start_day cannot be after end_day.")

    end_exclusive = end_day + timedelta(days=1)
    cursor = start_day
    windows: list[Window] = []

    while cursor < end_exclusive:
        if cadence == "monthly":
            next_boundary = _next_month_start(cursor)
        else:
            next_boundary = cursor + timedelta(days=14)

        window_end = min(next_boundary, end_exclusive)
        windows.append(Window(start=_at_midnight(cursor), end=_at_midnight(window_end)))
        cursor = window_end

    return windows


def build_manual_window(start_day: date, end_day: date) -> Window:
    if start_day > end_day:
        raise ValueError("start_day cannot be after end_day.")
    return Window(start=_at_midnight(start_day), end=_at_midnight(end_day + timedelta(days=1)))


def apply_window_limit(windows: list[Window], max_windows: int | None) -> list[Window]:
    if max_windows is None:
        return windows
    if max_windows < 1:
        raise ValueError("max_windows must be >= 1.")
    return windows[:max_windows]


def build_received_filter(window: Window) -> str:
    start_text = window.start.strftime("%m/%d/%Y %I:%M %p")
    end_text = window.end.strftime("%m/%d/%Y %I:%M %p")
    return f"[ReceivedTime] >= '{start_text}' AND [ReceivedTime] < '{end_text}'"


def _at_midnight(day: date) -> datetime:
    return datetime.combine(day, time.min)


def _next_month_start(day: date) -> date:
    if day.month == 12:
        return date(day.year + 1, 1, 1)
    return date(day.year, day.month + 1, 1)
