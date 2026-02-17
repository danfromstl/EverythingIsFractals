"""Text normalization helpers shared by notebook-extracted scripts."""

from __future__ import annotations

import re
from typing import Iterable

import pandas as pd


def clean_text(text: object, keep_commas: bool = False) -> str:
    """Normalize a title string to match notebook cleaning behavior."""
    if text is None:
        return ""

    normalized = str(text).lower()
    pattern = r"[^a-z0-9,\s]" if keep_commas else r"[^a-z0-9\s]"
    normalized = re.sub(pattern, "", normalized)
    normalized = re.sub(r"\s+", " ", normalized).strip()
    return normalized


def clean_text_series(series: pd.Series, keep_commas: bool = False) -> pd.Series:
    """Clean an entire pandas Series of text values."""
    return series.map(lambda value: clean_text(value, keep_commas=keep_commas))


def clean_text_list(values: Iterable[object], keep_commas: bool = False) -> list[str]:
    """Clean an iterable of titles for inference scripts."""
    return [clean_text(value, keep_commas=keep_commas) for value in values]

