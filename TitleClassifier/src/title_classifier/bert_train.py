"""BERT training pipeline extracted from v3-v6 notebook lineages."""

from __future__ import annotations

from dataclasses import asdict, dataclass
import json
from pathlib import Path
import random
from typing import Any

import joblib
import numpy as np
import torch
from torch.optim import AdamW
from torch.utils.data import DataLoader, RandomSampler, SequentialSampler, TensorDataset
from transformers import BertForSequenceClassification, BertTokenizer

from .data import encode_labels, load_raw_dataframe, prepare_classification_dataframe, split_train_val_test


@dataclass
class BertTrainingConfig:
    csv_path: str
    text_column: str = "SOC18_DMT"
    label_column: str = "SOC18_DOC"
    output_dir: str = "TitleClassifier/models/bert"
    model_name: str = "bert-base-uncased"
    holdout_size: float = 0.2
    val_fraction_of_holdout: float = 0.5
    random_state: int = 42
    batch_size: int = 32
    learning_rate: float = 2e-5
    epochs: int = 3
    max_length: int = 128
    keep_commas: bool = True
    stratify: bool = True
    device: str = "auto"


def train_bert_classifier(config: BertTrainingConfig) -> dict[str, Any]:
    """Fine-tune a BERT classifier and persist model + metadata artifacts."""
    set_seed(config.random_state)

    raw_df = load_raw_dataframe(config.csv_path)
    df = prepare_classification_dataframe(
        raw_df,
        text_column=config.text_column,
        label_column=config.label_column,
        keep_commas=config.keep_commas,
    )
    encoded_labels, encoder = encode_labels(df[config.label_column].values)

    split = split_train_val_test(
        texts=df[config.text_column].tolist(),
        labels=encoded_labels,
        holdout_size=config.holdout_size,
        val_fraction_of_holdout=config.val_fraction_of_holdout,
        random_state=config.random_state,
        stratify=config.stratify,
    )

    tokenizer = BertTokenizer.from_pretrained(config.model_name, use_fast=True)
    model = BertForSequenceClassification.from_pretrained(
        config.model_name,
        num_labels=len(encoder.classes_),
    )

    train_loader = _create_dataloader(
        tokenizer=tokenizer,
        texts=split.train_texts,
        labels=split.train_labels,
        batch_size=config.batch_size,
        max_length=config.max_length,
        training=True,
    )
    val_loader = _create_dataloader(
        tokenizer=tokenizer,
        texts=split.val_texts,
        labels=split.val_labels,
        batch_size=config.batch_size,
        max_length=config.max_length,
        training=False,
    )
    test_loader = _create_dataloader(
        tokenizer=tokenizer,
        texts=split.test_texts,
        labels=split.test_labels,
        batch_size=config.batch_size,
        max_length=config.max_length,
        training=False,
    )

    device = resolve_device(config.device)
    model.to(device)
    optimizer = AdamW(model.parameters(), lr=config.learning_rate)

    history: list[dict[str, float]] = []
    for epoch in range(config.epochs):
        train_loss = _train_one_epoch(model, train_loader, optimizer, device)
        val_metrics = evaluate_classifier(model, val_loader, device)
        history.append(
            {
                "epoch": float(epoch + 1),
                "train_loss": float(train_loss),
                "val_loss": float(val_metrics["loss"]),
                "val_accuracy": float(val_metrics["accuracy"]),
            }
        )

    test_metrics = evaluate_classifier(model, test_loader, device)

    output_dir = Path(config.output_dir)
    model_dir = output_dir / "model"
    output_dir.mkdir(parents=True, exist_ok=True)
    model_dir.mkdir(parents=True, exist_ok=True)

    model.save_pretrained(model_dir)
    tokenizer.save_pretrained(model_dir)
    encoder_path = output_dir / "label_encoder.joblib"
    joblib.dump(encoder, encoder_path)

    history_path = output_dir / "history.json"
    test_path = output_dir / "test_metrics.json"
    config_path = output_dir / "config.json"
    _write_json(history_path, {"epochs": history})
    _write_json(test_path, test_metrics)
    _write_json(config_path, asdict(config))

    return {
        "test_metrics": test_metrics,
        "history": history,
        "device": str(device),
        "dataset": {
            "rows_total": int(len(raw_df)),
            "rows_after_cleaning": int(len(df)),
            "classes": int(len(encoder.classes_)),
            "train_rows": int(len(split.train_texts)),
            "val_rows": int(len(split.val_texts)),
            "test_rows": int(len(split.test_texts)),
        },
        "artifacts": {
            "model_dir": str(model_dir),
            "label_encoder": str(encoder_path),
            "history": str(history_path),
            "test_metrics": str(test_path),
            "config": str(config_path),
        },
    }


def resolve_device(device_choice: str) -> torch.device:
    normalized = device_choice.lower().strip()
    if normalized == "auto":
        return torch.device("cuda" if torch.cuda.is_available() else "cpu")
    if normalized == "cuda":
        if not torch.cuda.is_available():
            raise RuntimeError("CUDA requested but no CUDA device is available.")
        return torch.device("cuda")
    if normalized == "cpu":
        return torch.device("cpu")
    raise ValueError("device must be one of: auto, cuda, cpu")


def set_seed(seed: int) -> None:
    random.seed(seed)
    np.random.seed(seed)
    torch.manual_seed(seed)
    if torch.cuda.is_available():
        torch.cuda.manual_seed_all(seed)


def evaluate_classifier(
    model: BertForSequenceClassification,
    dataloader: DataLoader,
    device: torch.device,
) -> dict[str, float]:
    model.eval()
    total_loss = 0.0
    total_correct = 0
    total_examples = 0

    with torch.no_grad():
        for batch in dataloader:
            input_ids, attention_mask, labels = [item.to(device) for item in batch]
            outputs = model(
                input_ids=input_ids,
                attention_mask=attention_mask,
                labels=labels,
            )
            total_loss += float(outputs.loss.item())
            predictions = torch.argmax(outputs.logits, dim=1)
            total_correct += int((predictions == labels).sum().item())
            total_examples += int(labels.size(0))

    avg_loss = total_loss / max(len(dataloader), 1)
    accuracy = total_correct / max(total_examples, 1)
    return {"loss": float(avg_loss), "accuracy": float(accuracy)}


def _train_one_epoch(
    model: BertForSequenceClassification,
    dataloader: DataLoader,
    optimizer: AdamW,
    device: torch.device,
) -> float:
    model.train()
    total_loss = 0.0

    for batch in dataloader:
        input_ids, attention_mask, labels = [item.to(device) for item in batch]
        optimizer.zero_grad()
        outputs = model(
            input_ids=input_ids,
            attention_mask=attention_mask,
            labels=labels,
        )
        loss = outputs.loss
        loss.backward()
        optimizer.step()
        total_loss += float(loss.item())

    return total_loss / max(len(dataloader), 1)


def _create_dataloader(
    tokenizer: BertTokenizer,
    texts: list[str],
    labels: np.ndarray,
    batch_size: int,
    max_length: int,
    training: bool,
) -> DataLoader:
    encodings = tokenizer(
        texts,
        padding=True,
        truncation=True,
        max_length=max_length,
        return_tensors="pt",
    )
    tensor_labels = torch.tensor(labels, dtype=torch.long)
    dataset = TensorDataset(encodings["input_ids"], encodings["attention_mask"], tensor_labels)
    sampler = RandomSampler(dataset) if training else SequentialSampler(dataset)
    return DataLoader(dataset, sampler=sampler, batch_size=batch_size)


def _write_json(path: Path, payload: dict[str, Any]) -> None:
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

