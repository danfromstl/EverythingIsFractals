"""CLI entrypoint for BERT fine-tuning on title datasets."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys


def _bootstrap_path() -> None:
    script_dir = Path(__file__).resolve().parent
    src_dir = script_dir.parent / "src"
    if str(src_dir) not in sys.path:
        sys.path.insert(0, str(src_dir))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Train BERT title classifier artifacts.")
    parser.add_argument("--csv-path", required=True, help="Input CSV path.")
    parser.add_argument("--text-column", default="SOC18_DMT", help="Text column name.")
    parser.add_argument("--label-column", default="SOC18_DOC", help="Label column name.")
    parser.add_argument("--output-dir", default="TitleClassifier/models/bert", help="Artifact output directory.")
    parser.add_argument("--model-name", default="bert-base-uncased", help="Base Hugging Face model name.")
    parser.add_argument("--holdout-size", type=float, default=0.2, help="Total holdout split fraction.")
    parser.add_argument(
        "--val-fraction-of-holdout",
        type=float,
        default=0.5,
        help="Validation fraction within holdout split.",
    )
    parser.add_argument("--random-state", type=int, default=42, help="Random seed.")
    parser.add_argument("--batch-size", type=int, default=32, help="Batch size.")
    parser.add_argument("--learning-rate", type=float, default=2e-5, help="AdamW learning rate.")
    parser.add_argument("--epochs", type=int, default=3, help="Training epochs.")
    parser.add_argument("--max-length", type=int, default=128, help="Tokenizer max sequence length.")
    parser.add_argument("--keep-commas", action="store_true", help="Keep commas during text cleaning.")
    parser.add_argument("--device", choices=["auto", "cpu", "cuda"], default="auto", help="Training device.")
    parser.add_argument(
        "--no-stratify",
        action="store_true",
        help="Disable stratified splitting even when possible.",
    )
    return parser.parse_args()


def main() -> None:
    _bootstrap_path()
    args = parse_args()

    from title_classifier.bert_train import BertTrainingConfig, train_bert_classifier

    config = BertTrainingConfig(
        csv_path=args.csv_path,
        text_column=args.text_column,
        label_column=args.label_column,
        output_dir=args.output_dir,
        model_name=args.model_name,
        holdout_size=args.holdout_size,
        val_fraction_of_holdout=args.val_fraction_of_holdout,
        random_state=args.random_state,
        batch_size=args.batch_size,
        learning_rate=args.learning_rate,
        epochs=args.epochs,
        max_length=args.max_length,
        keep_commas=args.keep_commas,
        stratify=not args.no_stratify,
        device=args.device,
    )
    result = train_bert_classifier(config)
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
