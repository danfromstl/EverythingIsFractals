# Everything Is Fractals

This repo currently contains three distinct experiment groups.

## 1) Title Classifier

### Notebook lineage

| Canonical | Notebook | Core approach | Suggested tag |
|---|---|---|---|
| Classical-1 | `TitleClassifier_v1_complete.ipynb` | TF-IDF + KMeans + logistic baseline | `tc-classic-1.0` |
| Classical-2 | `TitleClassifier_v2.ipynb` | Classical iteration (chunking/logging tweaks) | `tc-classic-1.1` |
| BERT-Proto | `TitleClassifier_v3_BERT.ipynb` | First BERT transition (`bert-base-uncased`) | `tc-bert-0.1` |
| BERT-DOC-Refactor | `TitleClassifier_v4_BERT_v2.ipynb` | DOC/DMT dataset refactor, larger label space | `tc-bert-1.0-rc1` |
| BERT-DOC-GPU | `TitleClassifier_v5_BERT_v3_HURRICANE.ipynb` | GPU-focused train/val flow on DOC labels | `tc-bert-1.0` |
| BERT-MGC-GPU | `TitleClassifier_v6_BERT_v3_HURRICANE2.ipynb` | GPU-focused variant on MGC labels | `tc-bert-1.1-mgc98` |

SpaceLaser side branch:

- `TitleClassifier/notebooks/spacelaser/SpaceLaser_v1_BERT_v4_HURRICANE.ipynb`
- `TitleClassifier/notebooks/spacelaser/SpaceLaser_v9_PretrainingOnTheSOC.ipynb`

### Scripted pipeline (extracted from notebooks)

- Package: `TitleClassifier/src/title_classifier/`
- CLI scripts: `TitleClassifier/scripts/`
- Dependencies: `TitleClassifier/requirements.txt`

Install dependencies:

```bash
pip install -r TitleClassifier/requirements.txt
```

## 2) AD Extraction + ID Transformations (VBA)

VBA modules now live in `ADExtraction/`:

- `ADExtraction/AD_Export.bas`
- `ADExtraction/ID_Transforms.bas`
- Legacy `randomNotes.bas` has been retired.

### AD export module

- Entry point: `ExportADUsersToCSV`
- Primary output: `allUsers_allSubdomains.csv`
- Manager/direct report output: `ManagersAndDRs.csv`
- Uses LDAP (`ADsDSOObject`) and writes pipe-delimited exports.

### ID transform versions

The transform module includes several generations of ID/checksum/obfuscation functions:

- Base checksums/hashes:
  - `SimpleHash`
  - `SimpleHashFormula`
  - `SimpleChecksum`
  - `SimpleChecksumHash`
  - `SimpleChecksumHash_v2`
- Base64-based variants:
  - `Base64ChecksumHash` (24-bit)
  - `Base64_Hash_8` (48-bit / 8-char base64 body)
- Offset encode/decode variants:
  - `OffsetEncode`, `OffsetEncode_v2`, `OffsetEncode_v3`, `OffsetEncode_v5`
  - `OffsetEncode_v6` / `OffsetDecode_v6`
  - `OffsetEncode_v7` / `OffsetDecode_v7`

These are practical anonymization/checksum transforms, not cryptographic hashes.

## 3) URL Parser (Darknet Diaries CSV pass)

Location:

- Script: `URLParser/URLparser.py`
- Data file: `URLParser/DarknetDiaries-CVS_Export.csv`

What it does now:

- Reads the local CSV export (not live RSS scraping at runtime).
- Scans episode descriptions for URL-like strings (`.com`, `.org`, `.tv`).
- Counts total discovered URLs and Twitter/X profile hits.

Dataset notes (current file):

- 114 episodes
- Columns: `Number, Title, URL, Duration, Duration (Minutes), Publish Date, Description`
- Date range in file: September 1, 2017 to April 5, 2022

---

Repo hygiene notes:

- `TitleClassifier/cmder/` is ignored in `.gitignore`.
- Notebook checkpoints are ignored via `**/.ipynb_checkpoints/`.
