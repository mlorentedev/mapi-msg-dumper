from datetime import date

import pytest

from mapi_msg_dumper.core.planning import apply_window_limit, build_auto_windows, build_received_filter


def test_build_auto_windows_monthly_partial_first_month() -> None:
    windows = build_auto_windows(start_day=date(2024, 1, 15), end_day=date(2024, 3, 2), cadence="monthly")

    assert len(windows) == 3
    assert windows[0].start.date() == date(2024, 1, 15)
    assert windows[0].end.date() == date(2024, 2, 1)
    assert windows[1].start.date() == date(2024, 2, 1)
    assert windows[1].end.date() == date(2024, 3, 1)
    assert windows[2].start.date() == date(2024, 3, 1)
    assert windows[2].end.date() == date(2024, 3, 3)


def test_build_auto_windows_biweekly() -> None:
    windows = build_auto_windows(start_day=date(2024, 1, 1), end_day=date(2024, 1, 31), cadence="biweekly")

    assert [window.start.date() for window in windows] == [date(2024, 1, 1), date(2024, 1, 15), date(2024, 1, 29)]
    assert [window.end.date() for window in windows] == [date(2024, 1, 15), date(2024, 1, 29), date(2024, 2, 1)]


def test_build_received_filter_uses_expected_outlook_format() -> None:
    window = build_auto_windows(start_day=date(2024, 4, 1), end_day=date(2024, 4, 2), cadence="monthly")[0]
    clause = build_received_filter(window)

    assert "[ReceivedTime] >=" in clause
    assert "[ReceivedTime] <" in clause
    assert "04/01/2024 12:00 AM" in clause


def test_apply_window_limit_reduces_batches() -> None:
    windows = build_auto_windows(start_day=date(2024, 1, 1), end_day=date(2024, 3, 31), cadence="monthly")
    limited = apply_window_limit(windows, 2)

    assert len(limited) == 2
    assert limited[0] == windows[0]
    assert limited[1] == windows[1]


def test_apply_window_limit_rejects_invalid_value() -> None:
    windows = build_auto_windows(start_day=date(2024, 1, 1), end_day=date(2024, 1, 31), cadence="monthly")

    with pytest.raises(ValueError):
        apply_window_limit(windows, 0)
