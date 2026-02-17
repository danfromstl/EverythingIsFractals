"""Dataset loading, cleaning, label encoding, and split helpers."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Sequence

import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import LabelEncoder

from .preprocess import clean_text_series


@dataclass
class SplitData:
    train_texts: list[str]
    train_labels: np.ndarray
    val_texts: list[str]
    val_labels: np.ndarray
    test_texts: list[str]
    test_labels: np.ndarray


def load_raw_dataframe(csv_path: str | Path) -> pd.DataFrame:
    """Load CSV data from disk."""
    path = Path(csv_path)
    if not path.exists():
        raise FileNotFoundError(f"CSV path not found: {path}")
    return pd.read_csv(path)


def prepare_classification_dataframe(
    df: pd.DataFrame,
    text_column: str,
    label_column: str,
    keep_commas: bool = False,
) -> pd.DataFrame:
    """Drop invalid rows and normalize text columns for classification tasks."""
    _validate_columns(df, [text_column, label_column])

    prepared = df.dropna(subset=[text_column, label_column]).copy()
    prepared[text_column] = clean_text_series(prepared[text_column], keep_commas=keep_commas)
    prepared = prepared[prepared[text_column].str.len() > 0].copy()
    return prepared


def encode_labels(raw_labels: Sequence[object]) -> tuple[np.ndarray, LabelEncoder]:
    """Label-encode target values and return both encoded labels and encoder."""
    encoder = LabelEncoder()
    encoded = encoder.fit_transform(list(raw_labels))
    return encoded, encoder


def can_stratify(labels: Sequence[int]) -> bool:
    """Return True when labels support stratified splitting."""
    values, counts = np.unique(np.asarray(labels), return_counts=True)
    return len(values) > 1 and int(counts.min()) >= 2


def split_train_val_test(
    texts: Sequence[str],
    labels: Sequence[int],
    holdout_size: float = 0.2,
    val_fraction_of_holdout: float = 0.5,
    random_state: int = 42,
    stratify: bool = True,
) -> SplitData:
    """Create 80/10/10-style splits (train/val/test by default)."""
    if not (0.0 < holdout_size < 1.0):
        raise ValueError("holdout_size must be between 0 and 1.")
    if not (0.0 < val_fraction_of_holdout < 1.0):
        raise ValueError("val_fraction_of_holdout must be between 0 and 1.")

    labels_array = np.asarray(labels)
    first_stratify = labels_array if stratify and can_stratify(labels_array) else None

    train_texts, holdout_texts, train_labels, holdout_labels = _safe_train_test_split(
        list(texts),
        labels_array,
        test_size=holdout_size,
        random_state=random_state,
        stratify=first_stratify,
    )

    second_stratify = holdout_labels if stratify and can_stratify(holdout_labels) else None
    val_texts, test_texts, val_labels, test_labels = _safe_train_test_split(
        holdout_texts,
        holdout_labels,
        test_size=(1.0 - val_fraction_of_holdout),
        random_state=random_state,
        stratify=second_stratify,
    )

    return SplitData(
        train_texts=list(train_texts),
        train_labels=np.asarray(train_labels),
        val_texts=list(val_texts),
        val_labels=np.asarray(val_labels),
        test_texts=list(test_texts),
        test_labels=np.asarray(test_labels),
    )


def _validate_columns(df: pd.DataFrame, required_columns: Sequence[str]) -> None:
    missing = [column for column in required_columns if column not in df.columns]
    if missing:
        raise KeyError(f"Missing required columns: {missing}")


def _safe_train_test_split(
    x: Sequence[object],
    y: Sequence[int],
    test_size: float,
    random_state: int,
    stratify: Sequence[int] | None,
):
    try:
        return train_test_split(
            x,
            y,
            test_size=test_size,
            random_state=random_state,
            stratify=stratify,
        )
    except ValueError:
        return train_test_split(
            x,
            y,
            test_size=test_size,
            random_state=random_state,
            stratify=None,
        )

