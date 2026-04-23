# New Mapper Project

A semantic-similarity mapper that takes a set of source items (events, issues,
findings, etc.) and suggests which L2 taxonomy categories each belongs to —
one, multiple, or none.

Derived from the ORE / PRSA / RAP mappers in the `risk_taxonomy_transformer`
project. See `HANDOVER.md` for the patterns and conventions carried forward.

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
│   └── mapper_config.yaml        # column names, thresholds, spaCy model
├── data/
│   ├── input/                    # put source + taxonomy xlsx here
│   └── output/                   # timestamped mapping_*.xlsx lands here
├── requirements.txt
├── .gitignore
├── README.md                     # this file
└── HANDOVER.md                   # context from the prior project
```

## Run

1. Put your taxonomy file in `data/input/` (default name: `Taxonomy.xlsx`). It
   needs at minimum `L1`, `L2`, `L2 Definition` columns. Optional: `L3`,
   `L3 Definition`, `L4`, `L4 Definition` — if present, text from those columns
   is folded into each L2's semantic vector for richer matching.
2. Put your source-item file in `data/input/` matching the glob in the yaml
   (default: `source_*.xlsx`). Required columns per the yaml.
3. Run:
   ```bash
   python mapper.py
   ```
4. Output lands in `data/output/mapping_<timestamp>.xlsx` with five sheets:
   - **All Mappings** — one row per item, reviewer-friendly.
   - **Needs Review** — side-by-side comparison workspace for ambiguous items
     (tool listed multiple candidate L2s; reviewer confirms).
   - **Summary** — counts + confidence distribution.
   - **L2 Distribution** — item count per L2 (exploded for multi-L2 rows).
   - **Raw Scores** — hidden by default, for threshold tuning.

## Configuration

Everything user-facing lives in `config/mapper_config.yaml`:

- **Column names** for both source and taxonomy files (match your enterprise
  export headers exactly).
- **spaCy model** — `en_core_web_lg` recommended, `en_core_web_md` as fallback.
- **Thresholds** — `min_similarity_score` (default 0.50), `high_similarity_score`
  (default 0.75), `ambiguity_margin_threshold` (null = auto-tune from data).
- **Closed statuses** — items with these statuses are excluded from mapping.

## Mapping classifications

| Status | What it means |
|---|---|
| **Suggested Match** | Top match is above the minimum threshold AND the margin to Match 2 is above the ambiguity threshold. Tool picked one (or a narrow cluster). |
| **Needs Review** | Top match is valid but margin to Match 2 is too small to rank confidently. Tool emits 1–3 candidate L2s; reviewer confirms. |
| **No Match** | Top match is below the minimum threshold. Nothing fit well enough. |

| Confidence | Meaning |
|---|---|
| Strong | Top score >= `high_similarity_score` |
| Moderate | Top score below `high_similarity_score` but valid |
| Review Required | Needs Review row |
| Weak | No Match row |
