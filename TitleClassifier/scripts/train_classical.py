"""CLI entrypoint for TF-IDF + logistic training."""

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
    parser = argparse.ArgumentParser(description="Train classical title classifier artifacts.")
    parser.add_argument("--csv-path", required=True, help="Input CSV path.")
    parser.add_argument("--text-column", default="organizationalPerson.title", help="Text column name.")
    parser.add_argument("--label-column", default="label", help="Label column name.")
    parser.add_argument("--output-dir", default="TitleClassifier/models/classical", help="Artifact output directory.")
    parser.add_argument("--test-size", type=float, default=0.2, help="Test split size fraction.")
    parser.add_argument("--random-state", type=int, default=42, help="Random seed.")
    parser.add_argument("--ngram-min", type=int, default=1, help="TF-IDF ngram min.")
    parser.add_argument("--ngram-max", type=int, default=2, help="TF-IDF ngram max.")
    parser.add_argument("--max-features", type=int, default=5000, help="TF-IDF max features.")
    parser.add_argument("--max-df", type=float, default=0.85, help="TF-IDF max_df.")
    parser.add_argument("--min-df", type=float, default=0.01, help="TF-IDF min_df.")
    parser.add_argument("--logistic-max-iter", type=int, default=1000, help="LogisticRegression max_iter.")
    parser.add_argument("--logistic-n-jobs", type=int, default=None, help="LogisticRegression n_jobs.")
    parser.add_argument("--keep-commas", action="store_true", help="Keep commas during text cleaning.")
    parser.add_argument("--run-kmeans-preview", action="store_true", help="Run KMeans inertia preview.")
    parser.add_argument("--kmeans-clusters", type=int, default=5, help="KMeans cluster count if preview enabled.")
    return parser.parse_args()


def main() -> None:
    _bootstrap_path()
    args = parse_args()

    from title_classifier.classical_train import ClassicalTrainingConfig, train_classical_classifier

    config = ClassicalTrainingConfig(
        csv_path=args.csv_path,
        text_column=args.text_column,
        label_column=args.label_column,
        output_dir=args.output_dir,
        test_size=args.test_size,
        random_state=args.random_state,
        ngram_min=args.ngram_min,
        ngram_max=args.ngram_max,
        max_features=args.max_features,
        max_df=args.max_df,
        min_df=args.min_df,
        logistic_max_iter=args.logistic_max_iter,
        logistic_n_jobs=args.logistic_n_jobs,
        keep_commas=args.keep_commas,
        run_kmeans_preview=args.run_kmeans_preview,
        kmeans_clusters=args.kmeans_clusters,
    )
    result = train_classical_classifier(config)
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
