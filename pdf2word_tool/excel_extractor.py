from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterable


TABLE_STRATEGIES: tuple[tuple[str, dict[str, Any]], ...] = (
    ("lines-strict", {"vertical_strategy": "lines_strict", "horizontal_strategy": "lines_strict", "snap_tolerance": 3, "join_tolerance": 3, "intersection_tolerance": 4, "edge_min_length": 12}),
    ("lines-tolerant", {"vertical_strategy": "lines", "horizontal_strategy": "lines", "snap_tolerance": 5, "join_tolerance": 8, "intersection_tolerance": 8, "edge_min_length": 12}),
    ("text", {"vertical_strategy": "text", "horizontal_strategy": "text", "min_words_vertical": 2, "min_words_horizontal": 1, "text_x_tolerance": 3, "text_y_tolerance": 3}),
)


@dataclass(frozen=True)
class TableExtraction:
    tables: list[list[list[str | None]]]
    strategy: str
    score: float
    stamp_objects_removed: int = 0


def _is_red(value: Any) -> bool:
    if not isinstance(value, (tuple, list)) or len(value) < 3:
        return False
    try:
        red, green, blue = (float(value[0]), float(value[1]), float(value[2]))
    except (TypeError, ValueError):
        return False
    scale = 255.0 if max(red, green, blue) > 1.0 else 1.0
    return red >= 0.55 * scale and red >= green * 1.45 and red >= blue * 1.45


def is_probable_stamp_object(obj: dict[str, Any]) -> bool:
    """Identify common red seal vector/text objects."""
    return _is_red(obj.get("stroking_color")) or _is_red(obj.get("non_stroking_color"))


def _table_score(tables: Iterable[list[list[str | None]]]) -> float:
    score = 0.0
    for table in tables:
        if not table:
            continue
        widths = [len(row) for row in table if row]
        if not widths:
            continue
        columns = max(widths)
        cells = len(table) * columns
        nonempty_rows = sum(1 for row in table if any(value is not None and str(value).strip() for value in row))
        filled = sum(1 for row in table for value in row if value is not None and str(value).strip())
        density = filled / cells
        structure_bonus = min(nonempty_rows, 20) * min(columns, 12) * 0.35
        empty_row_penalty = (len(table) - nonempty_rows) * 4.0
        fragment_penalty = 8.0 if nonempty_rows == 1 or columns == 1 else 0.0
        score += filled * 2.0 + density * 4.0 + structure_bonus - empty_row_penalty - fragment_penalty
    return score


def extract_best_tables(page: Any) -> TableExtraction:
    """Try several models, including a red-seal-filtered page, and retain the best result."""
    page_variants: list[tuple[str, Any, int]] = [("original", page, 0)]
    objects = getattr(page, "objects", {}) or {}
    removed = sum(
        1
        for object_list in objects.values()
        for obj in object_list
        if isinstance(obj, dict) and is_probable_stamp_object(obj)
    )
    if removed and hasattr(page, "filter"):
        filtered = page.filter(lambda obj: not is_probable_stamp_object(obj))
        page_variants.insert(0, ("seal-filtered", filtered, removed))

    candidates: list[TableExtraction] = []
    for variant_name, variant, removed_count in page_variants:
        for strategy_name, settings in TABLE_STRATEGIES:
            try:
                tables = variant.extract_tables(table_settings=settings) or []
            except Exception:
                continue
            candidates.append(TableExtraction(tables, f"{variant_name}/{strategy_name}", _table_score(tables), removed_count))

    if not candidates:
        return TableExtraction([], "none", 0.0, removed)
    # Filtering wins ties because seal geometry/text is not business table data.
    return max(candidates, key=lambda item: (item.score, item.stamp_objects_removed, len(item.tables)))
