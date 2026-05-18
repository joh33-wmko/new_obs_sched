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


def _normalize_semester_token(stem):
    match = re.search(r"(?<!\d)(\d{2}|\d{4})\s*([abAB])(?![A-Za-z0-9])", stem)
    if not match:
        return None

    year = match.group(1)
    semester = match.group(2).upper()
    if len(year) == 2:
        year = f"20{year}"

    return f"{year}{semester}"


def _normalize_staff_type(stem):
    match = re.search(r"(?<![A-Za-z0-9])(OA|NA|SA)(?![A-Za-z0-9])", stem, flags=re.IGNORECASE)
    if not match:
        return None

    return match.group(1).upper()


def _normalize_month_range(stem):
    month_pattern = (
        r"(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|"
        r"jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)"
    )
    match = re.search(rf"{month_pattern}\s*-\s*{month_pattern}", stem, flags=re.IGNORECASE)
    if not match:
        return None

    month_map = {
        "jan": "Jan",
        "feb": "Feb",
        "mar": "Mar",
        "apr": "Apr",
        "may": "May",
        "jun": "Jun",
        "jul": "Jul",
        "aug": "Aug",
        "sep": "Sep",
        "oct": "Oct",
        "nov": "Nov",
        "dec": "Dec",
    }

    start = month_map[match.group(1)[:3].lower()]
    end = month_map[match.group(2)[:3].lower()]
    return f"{start}-{end}"


def build_output_filename(file_path):
    stem = Path(file_path).stem
    semester = _normalize_semester_token(stem)
    staff_type = _normalize_staff_type(stem)
    month_range = _normalize_month_range(stem)

    if semester and staff_type:
        filename = f"{semester}_{staff_type}_Schedule"
        if month_range:
            filename = f"{filename}_{month_range}"
        return f"{filename}.csv"

    return f"{Path(file_path).stem}.csv"


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
