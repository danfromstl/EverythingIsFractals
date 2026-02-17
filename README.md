# Everything Is Fractals

This repo has multiple experiments. The section below makes the Title Classifier notebook history explicit and proposes a clean path to script-based Python modules.

## Title Classifier: Version Map

| Canonical | Notebook | What changed in this version | Suggested sub-version label |
|---|---|---|---|
| Classical-1 | `TitleClassifier_v1_complete.ipynb` | TF-IDF + KMeans + logistic regression baseline, with clustering/visualization steps. | `tc-classic-1.0` |
| Classical-2 | `TitleClassifier_v2.ipynb` | Same classical stack, plus chunked TF-IDF + progress/time logging. | `tc-classic-1.1` |
| BERT-Proto | `TitleClassifier_v3_BERT.ipynb` | First `bert-base-uncased` fine-tune path, `num_labels=6`, single dataloader flow. | `tc-bert-0.1` |
| BERT-DOC-Refactor | `TitleClassifier_v4_BERT_v2.ipynb` | Moves to `All_DOC-and-DMT_data.csv`, label encoding from `SOC18_DOC`, `num_labels=864`, stratified split setup appears. | `tc-bert-1.0-rc1` |
| BERT-DOC-GPU | `TitleClassifier_v5_BERT_v3_HURRICANE.ipynb` | Explicit train/val/test dataloaders, GPU device flow, validation loss tracking, save to `GPU_v1...pth`. | `tc-bert-1.0` |
| BERT-MGC-GPU | `TitleClassifier_v6_BERT_v3_HURRICANE2.ipynb` | Same loop style as v5, but dataset shifts to `All_MGC-and-DMT_data.csv`, labels from `SOC18_MGC`, `num_labels=98`. | `tc-bert-1.1-mgc98` |

## Helpful Sub-Version Notes

- `v2` likely needs a patch sub-version (`tc-classic-1.1.1`) because chunking currently uses `fit_transform` per chunk, which can create inconsistent feature spaces across chunks.
- `v4` looks like a transition build (`rc` style): split tensors are defined, but training still appears to iterate a combined dataloader path.
- `v5` and `v6` are better treated as dataset variants of the same trainer, not fully separate architecture generations.

## Extracted Python Scripts (Now Available)

Core package modules:

- `TitleClassifier/src/title_classifier/data.py`
- `TitleClassifier/src/title_classifier/preprocess.py`
- `TitleClassifier/src/title_classifier/bert_train.py`
- `TitleClassifier/src/title_classifier/classical_train.py`
- `TitleClassifier/src/title_classifier/predict.py`

CLI entry scripts:

- `TitleClassifier/scripts/train_bert.py`
- `TitleClassifier/scripts/train_classical.py`
- `TitleClassifier/scripts/predict_titles.py`
- `TitleClassifier/requirements.txt`

Install dependencies first:

```bash
pip install -r TitleClassifier/requirements.txt
```

Example usage:

```bash
python TitleClassifier/scripts/train_bert.py \
  --csv-path "C:/Offline_Storage/radiantClass/All_DOC-and-DMT_data.csv" \
  --text-column SOC18_DMT \
  --label-column SOC18_DOC \
  --output-dir TitleClassifier/models/bert_doc864 \
  --keep-commas \
  --epochs 3
```

```bash
python TitleClassifier/scripts/train_classical.py \
  --csv-path "C:/Offline_Storage/allTitles.csv" \
  --text-column organizationalPerson.title \
  --label-column label \
  --output-dir TitleClassifier/models/classical_v1
```

```bash
python TitleClassifier/scripts/predict_titles.py \
  --backend bert \
  --model-dir TitleClassifier/models/bert_doc864/model \
  --title "senior software engineer" \
  --title "director of operations"
```

## SpaceLaser Runtime Check

Based on the committed notebook contents:

- SpaceLaser notebooks now live under `TitleClassifier/notebooks/spacelaser/`.
- `SpaceLaser` notebooks show explicit CUDA/GPU logic and checks.
- No explicit Colab TPU hooks were found in-repo (`google.colab`, `torch_xla`, `xm.xla_device`, TPU runtime setup).
- It is still possible earlier/alternate local versions targeted Colab TPU, but the tracked versions here currently read as CUDA-oriented experiments.

## VBA Module Split

- Legacy `randomNotes.bas` content has been split into:
  - `AD_Export.bas` (Active Directory export workflow + progress logging)
  - `ID_Transforms.bas` (hash/checksum/offset-encode transform functions)
- `randomNotes.bas` is now kept as a migration pointer only.

## GitHub Language Visibility Tips

- `TitleClassifier/cmder/` is now ignored for this repo workflow so terminal-tooling files stay out of project history.
- Moving core logic into `.py` modules will make GitHub language distribution reflect Python work more clearly than notebook-only workflows.
- Optional: pair notebooks with scripts using Jupytext so notebook edits and plain Python stay synchronized.
