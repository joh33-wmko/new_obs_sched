"""Utilities for observing semester-related helpers.

This module is intentionally lightweight as a shared place for semester/date
helper functions used across scripts.
"""

import datetime
import re
from pathlib import Path


def semester_token_from_date(value):
    """Return semester token like '2026A' or '2026B' for a date-like value.

    Expects a ``datetime.date`` or ``datetime.datetime`` object.
    Semester A: Feb 1 - Jul 31
    Semester B: Aug 1 - Jan 31 (of the following year)
    """
    month = value.month
    year = value.year

    if 2 <= month <= 7:
        return f"{year}A"

    if month == 1:
        return f"{year - 1}B"

    return f"{year}B"


def update_config_sems(config_path=None, year=None):
    """Update SWOC semester ranges in config.live.ini.

    Args:
        config_path: Optional path to config file. Defaults to ./common/config.live.ini.
        year: Optional semester A year. Defaults to current year.

    Returns:
        dict with updated semester ranges.
    """
    target_year = int(year) if year is not None else datetime.date.today().year
    sem_a_start = f"{target_year:04d}-02-01"
    sem_a_end = f"{target_year:04d}-07-31"
    sem_b_start = f"{target_year:04d}-08-01"
    sem_b_end = f"{target_year + 1:04d}-01-31"

    default_path = Path(__file__).resolve().parent / "config.live.ini"
    cfg_path = Path(config_path) if config_path else default_path
    if not cfg_path.exists():
        raise FileNotFoundError(f"Config file not found: {cfg_path}")

    content = cfg_path.read_text(encoding="utf-8")

    sem_a_pattern = re.compile(
        r'("SEMESTER_A"\s*:\s*\{\s*"start_date"\s*:\s*)"[^"]+"(\s*,\s*"end_date"\s*:\s*)"[^"]+"',
        re.DOTALL,
    )
    sem_b_pattern = re.compile(
        r'("SEMESTER_B"\s*:\s*\{\s*"start_date"\s*:\s*)"[^"]+"(\s*,\s*"end_date"\s*:\s*)"[^"]+"',
        re.DOTALL,
    )

    new_content, sem_a_count = sem_a_pattern.subn(
        rf'\1"{sem_a_start}"\2"{sem_a_end}"',
        content,
        count=1,
    )
    new_content, sem_b_count = sem_b_pattern.subn(
        rf'\1"{sem_b_start}"\2"{sem_b_end}"',
        new_content,
        count=1,
    )

    if sem_a_count != 1 or sem_b_count != 1:
        raise ValueError("Could not locate SEMESTER_A/SEMESTER_B blocks in config file")

    cfg_path.write_text(new_content, encoding="utf-8")
    return {
        "config_path": str(cfg_path),
        "semester_a": {"start_date": sem_a_start, "end_date": sem_a_end},
        "semester_b": {"start_date": sem_b_start, "end_date": sem_b_end},
    }
