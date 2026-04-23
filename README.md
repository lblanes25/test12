# IAG Issues Backfill Mapper

Fills three taxonomy fields on pre-enterprise IAG issues — Issue Subtheme,
Root Cause Subtheme, and Risk Level 4 — by semantic similarity against
their respective taxonomies. Issue Theme is also filled as a by-product of
the chosen subtheme's parent.

For each blank target column on a row, the mapper picks the top-1 candidate
from the relevant taxonomy. Two of the three targets filter candidates by
a parent column already filled on the source row (Root Cause for J, Risk
Level 2 for L).

Derived from the ORE / PRSA / RAP mappers in the prior `risk_taxonomy_transformer`
project. See `HANDOVER.md` for context on patterns carried over and
deliberately stripped.

## Setup

```bash
python -m venv .venv
# Windows: .\.venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
python -m spacy download en_core_web_lg
```

## Layout

```
<project_root>/
├── mapper.py                     # main entry
├── config/
│   └── mapper_config.yaml        # source/taxonomy column names + targets
├── data/
│   ├── input/                    # put the IAG workbook here
│   └── output/                   # timestamped *_filled_*.xlsx lands here
├── requirements.txt
├── .gitignore
├── README.md                     # this file
└── HANDOVER.md                   # context from the prior project
```

## Run

1. Drop the IAG workbook into `data/input/` matching the glob in the yaml
   (default: `IAG Issues*.xlsx`). The same workbook holds the source sheet
   and the three taxonomy sheets.
2. Run:
   ```bash
   python mapper.py
   ```
3. Output lands in `data/output/<source-stem>_filled_<timestamp>.xlsx` — a
   copy of the source workbook with G/H/J/L filled where they were blank.
   Already-tagged rows are left untouched.

## Configuration

Everything user-facing lives in `config/mapper_config.yaml`:

- **`source.sheet_name`** — which tab in the workbook holds the issues.
- **`source.text_cols`** — list of column names whose text is concatenated
  to form each row's input for similarity scoring.
- **`source.explode_on`** — list of column names whose cells may contain
  multiple newline-separated values. Each such row is expanded into one
  row per value before mapping (cartesian product if more than one explode
  column has multiple values). Output workbook will have more rows than
  input. The same Issue ID will appear on multiple rows when its exploded
  column had multiple values.
- **`targets`** — one entry per column to fill. Each describes:
  - `target_col` — column on source sheet to fill (only when blank).
  - `taxonomy_sheet` / `taxonomy_parent_col` / `taxonomy_child_col` /
    `taxonomy_def_col` — where the candidates live and which columns to
    read.
  - `source_constraint_col` *(optional)* — column on source sheet whose
    value must equal `taxonomy_parent_col` for a row to be a candidate.
  - `parent_fill_col` *(optional)* — also write the chosen child's parent
    into this column on the source sheet (only when blank).
- **`spacy_model`** — `en_core_web_lg` recommended, `en_core_web_md` as
  fallback if disk/memory is tight.

## Behavior notes

- Top-1 fill only — no minimum-similarity threshold, no Needs Review
  classification, no candidate lists. The mapper picks the closest match
  and writes it.
- Existing tags are never overwritten. A row is touched only where the
  target column is blank.
- If a row's `source_constraint_col` is blank or its value isn't found in
  the taxonomy, that target is skipped for that row and logged.
- Run log written to `logs/mapper_log.txt`.
