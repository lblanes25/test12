"""
IAG Issues backfill mapper.

For each blank target column on the source sheet, picks the top-1 taxonomy
candidate by spaCy semantic similarity against the row's text columns and
writes it. Two of the three targets filter candidates by a parent column
already filled on the source.

When `source.explode_on` columns contain newline-separated values in a
cell, the row is expanded into one row per value (cartesian if multiple
explode columns) before mapping. All other cells are copied verbatim.
Useful when an Archer field permits multiple tags in a single cell —
each tag becomes its own row so the parent-constraint filter has a
single, unambiguous parent value to work with.

Config in `config/mapper_config.yaml`. Output: copy of the source workbook
with the source sheet rewritten (exploded + filled), written to
`data/output/<source-stem>_filled_<timestamp>.xlsx`.
"""

import logging
import shutil
from datetime import datetime
from itertools import product
from pathlib import Path

import numpy as np
import spacy
import yaml
from openpyxl import load_workbook

_PROJECT_ROOT = Path(__file__).parent
_LOG_DIR = _PROJECT_ROOT / "logs"
_LOG_DIR.mkdir(exist_ok=True)
_CONFIG_PATH = _PROJECT_ROOT / "config" / "mapper_config.yaml"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(_LOG_DIR / "mapper_log.txt", mode="w"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)


# =============================================================================
# Helpers
# =============================================================================

def _str(v) -> str:
    return "" if v is None else str(v).strip()


def _is_blank(s: str) -> bool:
    return s == "" or s.lower() in ("nan", "none")


def _norm(s: str) -> str:
    return s.strip().lower()


def _header_index(ws, header_name: str) -> int:
    for cell in ws[1]:
        if cell.value is not None and str(cell.value).strip() == header_name:
            return cell.column
    raise KeyError(
        f"Column {header_name!r} not found in sheet {ws.title!r}. "
        f"Headers: {[c.value for c in ws[1]]}"
    )


def _vectorize(text: str, nlp) -> np.ndarray:
    v = nlp(text).vector
    n = np.linalg.norm(v)
    return v / n if n > 0 else v


def _find_source_file(input_dir: Path, pattern: str) -> Path:
    files = sorted(input_dir.glob(pattern), key=lambda f: f.stat().st_mtime)
    if not files:
        raise FileNotFoundError(f"No file matching {pattern!r} in {input_dir}")
    return files[-1]


def _split_multiline(v) -> list[str]:
    s = _str(v)
    if not s:
        return [""]
    parts = [p.strip() for p in s.splitlines() if p.strip()]
    return parts if parts else [""]


def _explode_rows(rows: list[list], explode_col_indices: list[int]) -> list[tuple[int, list]]:
    """Cartesian-explode rows on multi-value cells in the given 1-based columns.

    Returns (original_row_index, exploded_row_data) tuples so warnings
    can reference the source row number.
    """
    out = []
    for orig_idx, row in enumerate(rows):
        value_lists = [_split_multiline(row[idx - 1]) for idx in explode_col_indices]
        if all(len(lst) <= 1 for lst in value_lists):
            out.append((orig_idx, list(row)))
            continue
        for combo in product(*value_lists):
            new_row = list(row)
            for i, idx in enumerate(explode_col_indices):
                new_row[idx - 1] = combo[i]
            out.append((orig_idx, new_row))
    return out


# =============================================================================
# Taxonomy loading
# =============================================================================

def _build_taxonomy(wb, target_cfg: dict, nlp) -> dict:
    """Vectorize one taxonomy sheet. Returns parent/child/vector arrays."""
    sheet = wb[target_cfg["taxonomy_sheet"]]
    parent_idx = _header_index(sheet, target_cfg["taxonomy_parent_col"])
    child_idx = _header_index(sheet, target_cfg["taxonomy_child_col"])
    def_idx = _header_index(sheet, target_cfg["taxonomy_def_col"])

    parents, children, vectors = [], [], []
    for r in range(2, sheet.max_row + 1):
        parent = _str(sheet.cell(r, parent_idx).value)
        child = _str(sheet.cell(r, child_idx).value)
        defn = _str(sheet.cell(r, def_idx).value)
        if _is_blank(child):
            continue
        text = ". ".join(p for p in (child, defn) if not _is_blank(p))
        parents.append(parent)
        children.append(child)
        vectors.append(_vectorize(text, nlp))

    return {
        "parents": parents,
        "parents_norm": [_norm(p) for p in parents],
        "children": children,
        "vectors": np.array(vectors),
    }


# =============================================================================
# Main
# =============================================================================

def main():
    with open(_CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f)

    src_cfg = cfg["source"]
    targets = cfg["targets"]
    spacy_model = cfg.get("spacy_model", "en_core_web_lg")

    input_dir = _PROJECT_ROOT / "data" / "input"
    output_dir = _PROJECT_ROOT / "data" / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    src_path = _find_source_file(input_dir, src_cfg["file_pattern"])
    logger.info(f"Source workbook: {src_path}")

    timestamp = datetime.now().strftime("%m%d%Y%I%M%p")
    out_path = output_dir / f"{src_path.stem}_filled_{timestamp}.xlsx"
    shutil.copy(src_path, out_path)
    wb = load_workbook(out_path)

    src_sheet = wb[src_cfg["sheet_name"]]
    n_cols = src_sheet.max_column
    n_data_rows = src_sheet.max_row - 1
    logger.info(f"Source sheet: {src_sheet.title!r} ({n_data_rows} data rows, {n_cols} columns)")

    logger.info(f"Loading spaCy: {spacy_model}")
    nlp = spacy.load(spacy_model)
    logger.info(f"  {len(nlp.vocab.vectors)} vectors, dim {nlp.vocab.vectors.shape[1]}")

    tax = {}
    for t in targets:
        tax[t["name"]] = _build_taxonomy(wb, t, nlp)
        logger.info(f"  {t['name']}: {len(tax[t['name']]['children'])} candidates")

    # Resolve column indices on the source sheet (1-based)
    id_idx = _header_index(src_sheet, src_cfg["id_col"])
    text_indices = [_header_index(src_sheet, c) for c in src_cfg["text_cols"]]
    target_idx = {t["name"]: _header_index(src_sheet, t["target_col"]) for t in targets}
    parent_fill_idx = {
        t["name"]: _header_index(src_sheet, t["parent_fill_col"])
        for t in targets if t.get("parent_fill_col")
    }
    constraint_idx = {
        t["name"]: _header_index(src_sheet, t["source_constraint_col"])
        for t in targets if t.get("source_constraint_col")
    }
    explode_indices = [_header_index(src_sheet, c) for c in src_cfg.get("explode_on", [])]

    # Read all data rows into memory, then explode
    source_rows = [
        [src_sheet.cell(r, c).value for c in range(1, n_cols + 1)]
        for r in range(2, src_sheet.max_row + 1)
    ]
    exploded = _explode_rows(source_rows, explode_indices)
    if explode_indices:
        added = len(exploded) - len(source_rows)
        logger.info(
            f"Row explosion on {src_cfg.get('explode_on')}: "
            f"{len(source_rows)} -> {len(exploded)} rows (+{added})"
        )

    fill_counts = {t["name"]: 0 for t in targets}
    skip_counts = {t["name"]: 0 for t in targets}
    parent_fill_counts = {n: 0 for n in parent_fill_idx}

    for orig_idx, row in exploded:
        # All values in `row` are 0-indexed; subtract 1 from sheet column index.
        if all(not _is_blank(_str(row[target_idx[t["name"]] - 1])) for t in targets):
            continue

        text_parts = [_str(row[i - 1]) for i in text_indices]
        text_parts = [p for p in text_parts if not _is_blank(p)]
        if not text_parts:
            issue_id = _str(row[id_idx - 1])
            logger.warning(
                f"Source row {orig_idx + 2} (Issue {issue_id!r}): no text in any of "
                f"{src_cfg['text_cols']} — skipped"
            )
            continue

        item_vec = _vectorize(". ".join(text_parts), nlp)

        for t in targets:
            name = t["name"]
            tcol_0 = target_idx[name] - 1
            if not _is_blank(_str(row[tcol_0])):
                continue

            entries = tax[name]
            mask = np.ones(len(entries["children"]), dtype=bool)

            if name in constraint_idx:
                src_parent = _str(row[constraint_idx[name] - 1])
                if _is_blank(src_parent):
                    issue_id = _str(row[id_idx - 1])
                    logger.warning(
                        f"Source row {orig_idx + 2} (Issue {issue_id!r}): {name!r} "
                        f"skipped — constraint column "
                        f"{t['source_constraint_col']!r} is blank"
                    )
                    skip_counts[name] += 1
                    continue
                target_parent_norm = _norm(src_parent)
                mask = np.array([p == target_parent_norm for p in entries["parents_norm"]])
                if not mask.any():
                    issue_id = _str(row[id_idx - 1])
                    logger.warning(
                        f"Source row {orig_idx + 2} (Issue {issue_id!r}): {name!r} "
                        f"skipped — no taxonomy entries with parent {src_parent!r}"
                    )
                    skip_counts[name] += 1
                    continue

            scores = entries["vectors"] @ item_vec
            scores = np.where(mask, scores, -np.inf)
            top = int(np.argmax(scores))

            row[tcol_0] = entries["children"][top]
            fill_counts[name] += 1

            # Fill parent only if currently blank — never overwrite.
            if name in parent_fill_idx:
                pcol_0 = parent_fill_idx[name] - 1
                chosen_parent = entries["parents"][top]
                if _is_blank(_str(row[pcol_0])) and not _is_blank(chosen_parent):
                    row[pcol_0] = chosen_parent
                    parent_fill_counts[name] += 1

    # Rewrite source sheet: clear data rows, write exploded+filled rows back.
    if n_data_rows > 0:
        src_sheet.delete_rows(2, n_data_rows)
    for i, (_, row) in enumerate(exploded):
        for j, val in enumerate(row, start=1):
            src_sheet.cell(row=i + 2, column=j, value=val)

    wb.save(out_path)

    logger.info("=" * 60)
    logger.info("FILL COMPLETE")
    logger.info(f"  Output rows: {len(exploded)} (from {len(source_rows)} source rows)")
    for t in targets:
        name = t["name"]
        line = f"  {name}: filled {fill_counts[name]}, skipped {skip_counts[name]}"
        if name in parent_fill_counts:
            line += f" (also filled {parent_fill_counts[name]} {t['parent_fill_col']!r})"
        logger.info(line)
    logger.info(f"  Output: {out_path}")
    logger.info("=" * 60)

    print(f"\nDone! Output: {out_path}")


if __name__ == "__main__":
    main()
