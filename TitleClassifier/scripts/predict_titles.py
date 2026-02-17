"""CLI entrypoint for title prediction using classical or BERT artifacts."""

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
    parser = argparse.ArgumentParser(description="Predict labels for title strings.")
    parser.add_argument("--backend", choices=["bert", "classical"], required=True, help="Inference backend.")
    parser.add_argument("--titles-file", help="Path to a plain-text file with one title per line.")
    parser.add_argument("--title", action="append", help="Single title value (repeatable).")

    parser.add_argument("--model-dir", help="BERT model directory (contains tokenizer + model files).")
    parser.add_argument("--model-path", help="Classical model joblib path.")
    parser.add_argument("--vectorizer-path", help="Classical TF-IDF vectorizer joblib path.")
    parser.add_argument("--encoder-path", help="Optional label encoder joblib path.")

    parser.add_argument("--batch-size", type=int, default=32, help="BERT inference batch size.")
    parser.add_argument("--max-length", type=int, default=128, help="BERT tokenizer max length.")
    parser.add_argument("--device", choices=["auto", "cpu", "cuda"], default="auto", help="BERT device.")
    parser.add_argument("--keep-commas", action="store_true", help="Keep commas during text cleaning.")
    return parser.parse_args()


def main() -> None:
    _bootstrap_path()
    args = parse_args()

    from title_classifier.predict import (
        load_titles_file,
        predict_with_bert,
        predict_with_classical,
    )

    titles = _resolve_titles(args)

    if args.backend == "bert":
        if not args.model_dir:
            raise SystemExit("--model-dir is required when --backend bert")
        predictions = predict_with_bert(
            titles=titles,
            model_dir=args.model_dir,
            encoder_path=args.encoder_path,
            batch_size=args.batch_size,
            max_length=args.max_length,
            device=args.device,
            keep_commas=args.keep_commas,
        )
    else:
        if not args.model_path or not args.vectorizer_path:
            raise SystemExit("--model-path and --vectorizer-path are required when --backend classical")
        predictions = predict_with_classical(
            titles=titles,
            model_path=args.model_path,
            vectorizer_path=args.vectorizer_path,
            encoder_path=args.encoder_path,
            keep_commas=args.keep_commas,
        )

    print(json.dumps(predictions, indent=2))


def _resolve_titles(args: argparse.Namespace) -> list[str]:
    titles: list[str] = []
    if args.titles_file:
        titles.extend(load_titles_file(Path(args.titles_file)))
    if args.title:
        titles.extend(args.title)
    if not titles:
        raise SystemExit("Provide at least one --title or --titles-file")
    return titles


if __name__ == "__main__":
    main()
