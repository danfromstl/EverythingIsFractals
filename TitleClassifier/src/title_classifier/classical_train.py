"""Classical NLP training pipeline extracted from v1/v2 notebooks."""

from __future__ import annotations

from dataclasses import asdict, dataclass
import json
from pathlib import Path
from typing import Any

import joblib
import numpy as np
from sklearn.cluster import KMeans
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import accuracy_score, precision_recall_fscore_support
from sklearn.model_selection import train_test_split

from .data import can_stratify, encode_labels, load_raw_dataframe, prepare_classification_dataframe


@dataclass
class ClassicalTrainingConfig:
    csv_path: str
    text_column: str = "organizationalPerson.title"
    label_column: str = "label"
    output_dir: str = "TitleClassifier/models/classical"
    test_size: float = 0.2
    random_state: int = 42
    ngram_min: int = 1
    ngram_max: int = 2
    max_features: int | None = 5000
    max_df: float = 0.85
    min_df: float = 0.01
    logistic_max_iter: int = 1000
    logistic_n_jobs: int | None = None
    keep_commas: bool = False
    run_kmeans_preview: bool = False
    kmeans_clusters: int = 5


def train_classical_classifier(config: ClassicalTrainingConfig) -> dict[str, Any]:
    """Train and persist TF-IDF + logistic baseline artifacts."""
    raw_df = load_raw_dataframe(config.csv_path)
    df = prepare_classification_dataframe(
        raw_df,
        text_column=config.text_column,
        label_column=config.label_column,
        keep_commas=config.keep_commas,
    )

    encoded_labels, encoder = encode_labels(df[config.label_column].values)
    texts = df[config.text_column].tolist()
    stratify_labels = encoded_labels if can_stratify(encoded_labels) else None

    x_train, x_test, y_train, y_test = train_test_split(
        texts,
        encoded_labels,
        test_size=config.test_size,
        random_state=config.random_state,
        stratify=stratify_labels,
    )

    vectorizer = TfidfVectorizer(
        ngram_range=(config.ngram_min, config.ngram_max),
        max_features=config.max_features,
        max_df=config.max_df,
        min_df=config.min_df,
    )

    x_train_vec = vectorizer.fit_transform(x_train)
    x_test_vec = vectorizer.transform(x_test)

    classifier = LogisticRegression(
        max_iter=config.logistic_max_iter,
        n_jobs=config.logistic_n_jobs,
    )
    classifier.fit(x_train_vec, y_train)
    predictions = classifier.predict(x_test_vec)

    precision, recall, f1, _ = precision_recall_fscore_support(
        y_test,
        predictions,
        average="macro",
        zero_division=0,
    )

    metrics = {
        "accuracy": float(accuracy_score(y_test, predictions)),
        "precision_macro": float(precision),
        "recall_macro": float(recall),
        "f1_macro": float(f1),
    }

    kmeans_preview: dict[str, Any] | None = None
    if config.run_kmeans_preview:
        tfidf_all = vectorizer.transform(texts)
        kmeans = KMeans(
            n_clusters=config.kmeans_clusters,
            n_init="auto",
            random_state=config.random_state,
        )
        kmeans.fit(tfidf_all)
        kmeans_preview = {
            "n_clusters": int(config.kmeans_clusters),
            "inertia": float(kmeans.inertia_),
        }

    output_dir = Path(config.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    model_path = output_dir / "logreg_model.joblib"
    vectorizer_path = output_dir / "tfidf_vectorizer.joblib"
    encoder_path = output_dir / "label_encoder.joblib"
    metrics_path = output_dir / "metrics.json"
    config_path = output_dir / "config.json"

    joblib.dump(classifier, model_path)
    joblib.dump(vectorizer, vectorizer_path)
    joblib.dump(encoder, encoder_path)
    _write_json(metrics_path, metrics)
    _write_json(config_path, asdict(config))

    result = {
        "metrics": metrics,
        "artifacts": {
            "model": str(model_path),
            "vectorizer": str(vectorizer_path),
            "label_encoder": str(encoder_path),
            "metrics": str(metrics_path),
            "config": str(config_path),
        },
        "dataset": {
            "rows_total": int(len(raw_df)),
            "rows_after_cleaning": int(len(df)),
            "classes": int(len(np.unique(encoded_labels))),
        },
    }
    if kmeans_preview is not None:
        result["kmeans_preview"] = kmeans_preview

    return result


def _write_json(path: Path, payload: dict[str, Any]) -> None:
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

