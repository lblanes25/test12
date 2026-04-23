# Handover from `risk_taxonomy_transformer`

Context and conventions carried forward from the project this mapper was
derived from. Point a fresh Claude Code session at this file on first
run — the patterns below are the ones worth preserving.

## User profile snapshot

- Audit / risk practitioner building tooling for audit-leader / risk-category-owner
  review. Not a career engineer — comfortable reading and editing Python + JS,
  not interested in large refactors or infra.
- Prefers Excel-like UX in HTML reports (column filters, hide/show cols, etc.).
- Values tight, scannable commits. Writes his own commits too.
- Demos drive scope. Feedback from real users (not speculation) wins over PM
  "wait until Phase 2" when it's a small, localized change.

## Working preferences

- **Column names go in config, never hardcoded.** All mapper column references
  live in `config/mapper_config.yaml`. Enterprise exports change headers;
  yaml is the seam.
- **Route work through project agents when available.** Prior project had
  `.claude/agents/` (project-manager, audit-leader, builder, QA). If the new
  project is substantial enough, recreate analogous agents keyed to its domain.
- **No archive to GitHub.** `archive/` is gitignored — don't stage from it.
- **No speculative scaffolds.** Don't pre-build ingestion for data formats
  that haven't arrived yet. Wait for the actual data.
- **No multi-line rationale comments in code.** Comments explain *why* (one
  line max). Design decisions belong in commit messages / PR descriptions,
  not in the source.
- **Demo feedback from real users can jump the queue.** Even if it's formally
  Phase 2 polish, if an actual user asked for it during a demo, it's validated
  scope — ship it.

## Semantic-mapping patterns worth keeping

### 1. Aggregate reference vectors by canonical name, not per-row

Enterprise taxonomy files are often at L4 grain (one row per L4 leaf) with
L2 + L2 Definition repeated across rows. If you build one vector per row, you
get 9+ duplicate vectors per L2, top-N ranking picks up duplicates with
identical scores, and everything lands in "Needs Review."

Fix: group by L2 name, fold L3/L4 text into each L2's semantic vector,
produce one unique vector per L2. This is what `build_reference_vectors()`
in `mapper.py` does. Keep this.

### 2. Auto-tune the ambiguity margin threshold from data

Don't hardcode a margin like `0.05`. spaCy cosine scores are compressed.
Sample P25 of the margin distribution from actual data and clamp to
`[0.01, 0.05]`. The `determine_ambiguity_threshold()` function does this.

### 3. Three statuses, with candidate lists on Needs Review

- **Suggested Match** — tool picked confidently.
- **Needs Review** — tool emits 1–3 candidate L2s (all above min threshold).
  Reviewer confirms. Downstream consumers still ingest all candidates.
- **No Match** — below threshold. Dropped.

Reviewer workload shrinks from "map every item" to "confirm primary from
a short candidate list." Don't force a single pick on genuinely ambiguous
cases.

### 4. spaCy model is configurable via yaml

Default to `en_core_web_lg` (larger, more accurate). Fall back to
`en_core_web_md` if disk/memory is tight. The yaml `spacy_model` setting
drives `spacy.load()`.

### 5. Filter rows at ingestion, not display time

If certain items are out of scope (closed statuses, excluded categories),
exclude them BEFORE vector building. They compete in top-N ranking
otherwise and can hijack real matches.

## Things to adapt for the new project

Stripped from the template (reintroduce if needed):

- **L3-split bucket logic** — in the prior project, External Fraud was split
  into First Party / Victim Fraud at L3, and we rerouted Internal Fraud
  to its own L2. That's taxonomy-specific. If your taxonomy has similar
  sub-category splits, reintroduce via a `_bucket_for(l2, l3)` function.
- **Alias normalization** (`_L2_ALIASES` in the prior project's
  `normalization.py`) — only needed if your taxonomy renames happen
  independently of your source data's column values. Skip until you have
  a specific rename to handle.
- **Filter to evaluated L2s** — prior project had a yaml `new_taxonomy`
  list of evaluated L2s; mapper filtered the reference df to only those
  so not-assessed L2s didn't compete in ranking. Add if you have the
  same split of "in the file but not for evaluation."

## Output conventions

- **Timestamp in filenames:** `mapping_MMDDYYYYHHMMA/PM.xlsx`. Use
  `datetime.now().strftime("%m%d%Y%I%M%p")`.
- **Multi-sheet workbook** — All Mappings (visible), Needs Review (visible),
  Summary (visible), L2 Distribution (visible), Raw Scores (hidden).
- **Conditional fills** for Mapping Status column (green / yellow / gray)
  and Match Confidence column (four bands).
- **Freeze panes** on All Mappings (column C, row 2) and Needs Review
  (row 2 only).
- **Wrap text** on description-column cells; fixed width on everything
  else.

## When you're ready to grow

- Add an HTML report for reviewers to click through candidates instead of
  editing Excel. The prior project's `export_html_report.py` has patterns
  for entity-scoped view, column filters, click-to-expand cells, etc.
  Reach back to that project for inspiration — don't try to port wholesale.
- Extract shared library (spaCy loading, vector building, classification)
  if you end up with 3+ mappers that duplicate the same core.
