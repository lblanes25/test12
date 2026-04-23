"""
IAG Issues backfill mapper.

For each blank target column on the source sheet, picks the top-1 taxonomy
candidate by spaCy semantic similarity against the row's text columns and
writes it. Two of the three targets filter candidates by a parent column
already filled on the source.

Config in `config/mapper_config.yaml` describes source sheet, text columns,
and the three targets (each with taxonomy sheet, parent/child/definition
columns, and an optional `source_constraint_col` for parent filtering).

Output: a copy of the source workbook with blanks filled, written to
`data/output/<source-stem>_filled_<timestamp>.xlsx`.
"""

import logging
import shutil
from datetime import datetime
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

def _cell(ws, row, col) -> str:
    v = ws.cell(row=row, column=col).value
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
        parent = _cell(sheet, r, parent_idx)
        child = _cell(sheet, r, child_idx)
        defn = _cell(sheet, r, def_idx)
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
    logger.info(f"Source sheet: {src_sheet.title!r} ({src_sheet.max_row - 1} data rows)")

    logger.info(f"Loading spaCy: {spacy_model}")
    nlp = spacy.load(spacy_model)
    logger.info(
        f"  {len(nlp.vocab.vectors)} vectors, dim {nlp.vocab.vectors.shape[1]}"
    )

    # Build per-target taxonomy vectors
    tax = {}
    for t in targets:
        tax[t["name"]] = _build_taxonomy(wb, t, nlp)
        logger.info(f"  {t['name']}: {len(tax[t['name']]['children'])} candidates")

    # Resolve source column indices once
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

    fill_counts = {t["name"]: 0 for t in targets}
    skip_counts = {t["name"]: 0 for t in targets}
    parent_fill_counts = {n: 0 for n in parent_fill_idx}

    for r in range(2, src_sheet.max_row + 1):
        # Skip rows where every target is already filled
        if all(not _is_blank(_cell(src_sheet, r, target_idx[t["name"]])) for t in targets):
            continue

        text_parts = [_cell(src_sheet, r, i) for i in text_indices]
        text_parts = [p for p in text_parts if not _is_blank(p)]
        if not text_parts:
            logger.warning(f"Row {r}: no text in any of {src_cfg['text_cols']} — skipped")
            continue

        item_vec = _vectorize(". ".join(text_parts), nlp)

        for t in targets:
            name = t["name"]
            if not _is_blank(_cell(src_sheet, r, target_idx[name])):
                continue

            entries = tax[name]
            mask = np.ones(len(entries["children"]), dtype=bool)

            if name in constraint_idx:
                src_parent = _cell(src_sheet, r, constraint_idx[name])
                if _is_blank(src_parent):
                    logger.warning(
                        f"Row {r}: {name!r} skipped — constraint column "
                        f"{t['source_constraint_col']!r} is blank"
                    )
                    skip_counts[name] += 1
                    continue
                target_parent_norm = _norm(src_parent)
                mask = np.array([p == target_parent_norm for p in entries["parents_norm"]])
                if not mask.any():
                    logger.warning(
                        f"Row {r}: {name!r} skipped — no taxonomy entries with "
                        f"parent {src_parent!r}"
                    )
                    skip_counts[name] += 1
                    continue

            scores = entries["vectors"] @ item_vec
            scores = np.where(mask, scores, -np.inf)
            top = int(np.argmax(scores))

            chosen_child = entries["children"][top]
            chosen_parent = entries["parents"][top]
            src_sheet.cell(row=r, column=target_idx[name], value=chosen_child)
            fill_counts[name] += 1

            # Fill parent only if currently blank — don't overwrite existing values.
            if name in parent_fill_idx:
                pcol = parent_fill_idx[name]
                if _is_blank(_cell(src_sheet, r, pcol)) and not _is_blank(chosen_parent):
                    src_sheet.cell(row=r, column=pcol, value=chosen_parent)
                    parent_fill_counts[name] += 1

    wb.save(out_path)

    logger.info("=" * 60)
    logger.info("FILL COMPLETE")
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
