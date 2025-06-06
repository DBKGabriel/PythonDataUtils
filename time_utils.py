#!/usr/bin/env python3
"""
Hopefully final version of UTC - ET time zone converter
- Pure std-lib when Python ≥ 3.9 (`zoneinfo`); otherwise falls back to `pytz`.
- Bidirectional conversion, smart epoch-unit detection, lazy annotations.
- CLI flags: --units {auto|ms|s}  --fmt FORMAT  --to-utc  --info
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime, timezone
from typing import Union

# Backend detection of time-zone
_BACKEND: str

try: # Figured I should use a more modern function. 
     # Keeping the pytz logic as a fallback though
    from zoneinfo import ZoneInfo
    UTC = timezone.utc
    ET = ZoneInfo("America/New_York")
    _BACKEND = "zoneinfo"
except ModuleNotFoundError:  # If it's an older python version (< 3.9)
    try:
        import pytz  # type: ignore
        UTC = pytz.UTC            # pylint: disable=no-member
        ET = pytz.timezone("America/New_York")  # type: ignore
        _BACKEND = "pytz"
    except ModuleNotFoundError as exc:  # no zoneinfo, no pytz
        raise ImportError(
            "Neither zoneinfo (Py≥3.9) nor pytz is available; "
            "install one of them to use time_utils."
        ) from exc



def _to_datetime_utc(
    ts: Union[int, float, datetime],
    *,
    units: str,
) -> datetime:
    """Turns any timestamp-ish looking thing into a UTC aware datetime"""
    # Lotta numbers? Hope its seconds or mils
    if isinstance(ts, (int, float)):
        if units == "auto":
            units = "ms" if ts >= 1_000_000_000_000 else "s"
        if units not in {"ms", "s"}:
            raise ValueError("units must be 'auto', 'ms', or 's'")

        seconds = ts / 1_000 if units == "ms" else ts
        return datetime.fromtimestamp(seconds, tz=UTC)

    # Datetime? Slap UTK on it if not otherwise specified.
    if isinstance(ts, datetime):
        return ts.replace(tzinfo=UTC) if ts.tzinfo is None else ts.astimezone(UTC)

    # This is what you get for trying to be lazy, Daniel
    raise TypeError(f"Unsupported type: {type(ts).__name__}")


# The actual util to be called
def utc_to_eastern(
    ts: Union[int, float, datetime],
    *,
    units: str = "auto",
) -> datetime:
    """Convert **UTC** epoch/datetime aware Eastern datetime."""
    return _to_datetime_utc(ts, units=units).astimezone(ET)


# Had to add the inverse conversion to keep it Thanos-approved
def eastern_to_utc( 
    ts: Union[int, float, datetime],
    *,
    units: str = "auto",
) -> datetime:
    """Convert **Eastern** epoch/datetime aware UTC datetime."""
    if isinstance(ts, (int, float)):
        # Same story as above. No one's ever going to read these comments.
        if units == "auto":
            units = "ms" if ts >= 1_000_000_000_000 else "s"
        seconds = ts / 1_000 if units == "ms" else ts
        dt_et = datetime.fromtimestamp(seconds, tz=ET)
    elif isinstance(ts, datetime):
        dt_et = ts.replace(tzinfo=ET) if ts.tzinfo is None else ts.astimezone(ET)
    else:
        raise TypeError(f"Unsupported type: {type(ts).__name__}")

    return dt_et.astimezone(UTC)


def format_timestamp(ts: datetime, fmt: str = "%Y-%m-%d %H:%M:%S %Z") -> str:
    """Why'd you pass junk? Return *ts* formatted with *fmt*; requires a `datetime`."""
    if not isinstance(ts, datetime):
        raise TypeError("ts must be a datetime object")
    return ts.strftime(fmt)


def get_backend_info() -> dict[str, str | bool]:
    """This exists because I got tired of debugging."""
    return {
        "backend": _BACKEND,
        "python": f"{sys.version_info.major}.{sys.version_info.minor}",
        "dst_aware": True,  # guaranteed if we reached here
        "utc_repr": str(UTC),
        "eastern_repr": str(ET),
    }


# I like poking my scripts from a shell. Give me cmd or give me death!
def _build_cli() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="time_utils",
        description="Convert timestamps between UTC and US Eastern.",
    )
    p.add_argument("timestamps", nargs="+",
                   help="Epoch seconds/millis or ISO-8601 datetime")
    p.add_argument("--units", choices=["auto", "ms", "s"], default="auto",
                   help="Numeric input units (default: auto ≥1e12→ms)")
    p.add_argument("--fmt", default="%Y-%m-%d %H:%M:%S %Z",
                   help="strftime pattern for output")
    p.add_argument("--to-utc", action="store_true",
                   help="Convert *from* Eastern *to* UTC (default is UTC→Eastern)")
    p.add_argument("--info", action="store_true",
                   help="Show backend diagnostics and exit")
    return p


def _parse_timestamp(raw: str) -> Union[int, float, datetime]:
    """We'll pretend any input is useful for now."""
    
    # 1st guess: is it a number?
    for cast in (int, float):
        try:
            return cast(raw)
        except ValueError:
            pass
    # 2nd guess: Did you use ISO? Karen's always talking about ISO
    try:
        return datetime.fromisoformat(raw.replace("Z", "+00:00"))
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"Unrecognised timestamp: {raw}") from exc


def main(argv: list[str] | None = None) -> None:  # CLI entry
    parser = _build_cli()
    args = parser.parse_args(argv)

    if args.info:
        for k, v in get_backend_info().items():
            print(f"{k}: {v}")
        return

    convert = eastern_to_utc if args.to_utc else utc_to_eastern
    direction = "ET to UTC" if args.to_utc else "UTC to ET"
    print(f"Converting ({direction}):")

    for raw in args.timestamps:
        try:
            ts_obj = _parse_timestamp(raw)
            out = convert(ts_obj, units=args.units)
            print(f"  {raw} → {format_timestamp(out, args.fmt)}") 
        except Exception as exc:  # noqa: BLE001  report per-item failure
            print(f"  {raw} → ERROR: {exc}") #Yes, I know this is a broad except. Linters, please don't complain at me


# This is going to be in a module at some point
if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit("\nInterrupted by user")
