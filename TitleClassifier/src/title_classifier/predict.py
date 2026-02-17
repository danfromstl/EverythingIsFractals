"""Inference helpers for both extracted classical and BERT pipelines."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Iterable

import joblib
import torch
from transformers import AutoModelForSequenceClassification, AutoTokenizer

from .bert_train import resolve_device
from .preprocess import clean_text_list


def predict_with_bert(
    titles: Iterable[object],
    model_dir: str | Path,
    encoder_path: str | Path | None = None,
    batch_size: int = 32,
    max_length: int = 128,
    device: str = "auto",
    keep_commas: bool = True,
) -> list[dict[str, Any]]:
    """Run predictions using a fine-tuned Hugging Face sequence classifier."""
    cleaned_titles = clean_text_list(titles, keep_commas=keep_commas)
    device_obj = resolve_device(device)
    path = Path(model_dir)

    tokenizer = AutoTokenizer.from_pretrained(path)
    model = AutoModelForSequenceClassification.from_pretrained(path)
    model.to(device_obj)
    model.eval()

    label_encoder = _load_optional_encoder(path, encoder_path)
    predictions: list[dict[str, Any]] = []

    for start in range(0, len(cleaned_titles), batch_size):
        batch_titles = cleaned_titles[start : start + batch_size]
        if not batch_titles:
            continue

        encoded = tokenizer(
            batch_titles,
            padding=True,
            truncation=True,
            max_length=max_length,
            return_tensors="pt",
        )
        encoded = {key: value.to(device_obj) for key, value in encoded.items()}

        with torch.no_grad():
            outputs = model(**encoded)
            probabilities = torch.softmax(outputs.logits, dim=1)
            confidence, class_ids = torch.max(probabilities, dim=1)

        for index, class_id in enumerate(class_ids.tolist()):
            raw_text = batch_titles[index]
            conf = float(confidence[index].item())
            decoded = _decode_label(class_id, label_encoder)
            predictions.append(
                {
                    "title": raw_text,
                    "class_id": int(class_id),
                    "label": decoded,
                    "confidence": conf,
                }
            )

    return predictions


def predict_with_classical(
    titles: Iterable[object],
    model_path: str | Path,
    vectorizer_path: str | Path,
    encoder_path: str | Path | None = None,
    keep_commas: bool = False,
) -> list[dict[str, Any]]:
    """Run predictions using TF-IDF + logistic baseline artifacts."""
    cleaned_titles = clean_text_list(titles, keep_commas=keep_commas)

    classifier = joblib.load(model_path)
    vectorizer = joblib.load(vectorizer_path)
    label_encoder = joblib.load(encoder_path) if encoder_path else None

    features = vectorizer.transform(cleaned_titles)
    class_ids = classifier.predict(features)

    if hasattr(classifier, "predict_proba"):
        probabilities = classifier.predict_proba(features)
        confidence_scores = probabilities.max(axis=1)
    else:
        confidence_scores = [None] * len(cleaned_titles)

    results: list[dict[str, Any]] = []
    for index, class_id in enumerate(class_ids.tolist()):
        confidence = confidence_scores[index]
        results.append(
            {
                "title": cleaned_titles[index],
                "class_id": int(class_id),
                "label": _decode_label(class_id, label_encoder),
                "confidence": float(confidence) if confidence is not None else None,
            }
        )
    return results


def load_titles_file(path: str | Path) -> list[str]:
    lines = Path(path).read_text(encoding="utf-8").splitlines()
    return [line.strip() for line in lines if line.strip()]


def _load_optional_encoder(model_dir: Path, encoder_path: str | Path | None):
    if encoder_path:
        return joblib.load(encoder_path)

    sibling_encoder = model_dir.parent / "label_encoder.joblib"
    if sibling_encoder.exists():
        return joblib.load(sibling_encoder)
    return None


def _decode_label(class_id: int, label_encoder) -> str | None:
    if label_encoder is None:
        return None
    return str(label_encoder.inverse_transform([class_id])[0])

