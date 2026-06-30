"""Resolve parsed locator specs against xlwings worksheets.

Six addressing strategies:
  1. Label-anchored relative
  2. Header + entry block (repeating rows)
  3. Named ranges (German Excel defined names)
  4. Absolute address
  5. Column + row-offset within a block
  6. Fixed result rows/cols
"""

from __future__ import annotations

import logging
import re
import unicodedata
from typing import Any

import xlwings as xw

logger = logging.getLogger(__name__)

_CELL_REF_RE = re.compile(r"^[A-Z]{1,3}\d+$")
_NAMED_RANGE_RE = re.compile(r"^[A-Z][A-Za-z0-9]*(_[A-Za-z0-9]+)+$")

SPARSE_ROW_BREAK_THRESHOLD = 3


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def norm(value: Any) -> str:
    """Normalize a label for comparison: NFKC, NBSP→space, strip, casefold."""
    if value is None:
        return ""
    s = str(value)
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("\xa0", " ")
    s = " ".join(s.split())
    return s.strip().casefold()


def prefer_si_sheet(sheet_name: str, available: list[str]) -> str:
    """Return the SI variant of a sheet name if it exists, otherwise the original."""
    si_name = f"{sheet_name} SI"
    if si_name in available:
        return si_name
    return sheet_name


def col_to_idx(col: str) -> int:
    """Convert column letters (A, ..., AA) to 1-based index."""
    result = 0
    for ch in col.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


def field_col(spec: str | dict) -> str:
    """Extract column letter from a field spec (string or dict with 'column')."""
    return spec if isinstance(spec, str) else spec.get("column", "A")


def cell_value(ws: xw.Sheet, col: str, row: int) -> Any:
    """Read a single cell by column letter(s) and row number."""
    return ws.range((row, col_to_idx(col))).value


def find_row_in_col(
    ws: xw.Sheet, col: str, needle: str, *,
    contains: bool = True, start_from: int = 1,
) -> int | None:
    """Return the first row where ``col``'s cell text matches *needle*.

    Reads the entire column in one AppleScript call for performance.
    """
    needle_n = norm(needle)
    if not needle_n:
        return None
    col_idx = col_to_idx(col)
    last_row = ws.used_range.last_cell.row
    if start_from > last_row:
        return None
    values = ws.range((start_from, col_idx), (last_row, col_idx)).value
    if values is None:
        return None
    if not isinstance(values, list):
        values = [values]
    for i, cell_val in enumerate(values):
        cell_n = norm(cell_val)
        if not cell_n:
            continue
        if contains:
            if needle_n in cell_n:
                return start_from + i
        else:
            if cell_n == needle_n:
                return start_from + i
    return None


def parse_cell_ref(ref: str) -> tuple[str, int]:
    """Split 'AB123' into ('AB', 123)."""
    m = re.match(r"^([A-Z]+)(\d+)$", ref)
    if not m:
        raise ValueError(f"Invalid cell reference: {ref!r}")
    return m.group(1), int(m.group(2))


def _is_cell_ref(s: str) -> bool:
    return bool(_CELL_REF_RE.match(s))


def _is_named_range(s: str) -> bool:
    return bool(_NAMED_RANGE_RE.match(s))


def is_header_row(row_data: dict[str, Any]) -> bool:
    """Return True if a block row looks like a header rather than data."""
    values = [v for k, v in row_data.items() if k != "_row" and v is not None]
    if not values:
        return False
    return all(isinstance(v, str) for v in values)


def is_entry_row_header(
    ws: xw.Sheet, row: int, column_fields: dict[str, dict],
) -> bool:
    """Check if the entry locator row is a column header, not a data row."""
    row_data = {name: cell_value(ws, field_col(spec), row)
                for name, spec in column_fields.items()}
    return is_header_row(row_data)


# ---------------------------------------------------------------------------
# Strategy 1: Label-anchored relative
# ---------------------------------------------------------------------------

def resolve_label_anchored(
    ws: xw.Sheet,
    locator_col: str,
    locator_string: str,
    input_col: str,
    row_offset: int = 0,
) -> Any:
    """Find locator_string in locator_col, read input_col."""
    row = find_row_in_col(ws, locator_col, locator_string)
    if row is None:
        logger.warning(
            "Label %r not found in column %s of sheet %r",
            locator_string, locator_col, ws.name,
        )
        return None
    return cell_value(ws, input_col, row + row_offset)


# ---------------------------------------------------------------------------
# Strategy 2: Header + entry block (repeating rows)
# ---------------------------------------------------------------------------

_DEFAULT_END_MARKER = "Unhide additional rows"


def resolve_block(
    ws: xw.Sheet,
    header_locator: dict,
    entry_locator: dict,
    column_fields: dict[str, dict],
    *,
    end_marker: str = _DEFAULT_END_MARKER,
    entry_row_start: int | None = None,
) -> list[dict[str, Any]]:
    """Iterate a repeating block, returning one dict per data row.

    Reads all needed columns in a single batch for performance.
    """
    entry_col = entry_locator.get("col", header_locator.get("col", "A"))

    if entry_row_start is not None:
        start_row = entry_row_start
    else:
        entry_string = entry_locator.get("string", "")
        if not entry_string:
            return []
        hdr_row = find_row_in_col(
            ws, header_locator["col"], header_locator["string"]
        )
        if hdr_row is None:
            logger.warning(
                "Block header %r not found in column %s",
                header_locator["string"], header_locator["col"],
            )
            return []
        start_row_found = find_row_in_col(
            ws, entry_col, entry_string, start_from=hdr_row,
        )
        if start_row_found is None:
            return []
        if is_entry_row_header(ws, start_row_found, column_fields):
            start_row = max(start_row_found + 1, hdr_row)
        else:
            start_row = max(start_row_found, hdr_row)

    last_row = ws.used_range.last_cell.row
    if start_row > last_row:
        return []

    # Build column index map: collect all unique columns we need to read
    entry_col_idx = col_to_idx(entry_col)
    field_names = list(column_fields.keys())
    field_col_idxs = [col_to_idx(field_col(column_fields[f])) for f in field_names]
    all_col_idxs = sorted(set(field_col_idxs + [entry_col_idx]))
    min_col = min(all_col_idxs)
    max_col = max(all_col_idxs)

    # One AppleScript call: read the entire rectangular region
    raw = ws.range(
        (start_row, min_col), (last_row, max_col)
    ).value
    if raw is None:
        return []
    # xlwings returns flat list for single-row OR single-column ranges
    if not isinstance(raw, list):
        raw = [[raw]]
    elif not isinstance(raw[0], list):
        if min_col == max_col:
            raw = [[v] for v in raw]
        else:
            raw = [raw]

    entry_offset = entry_col_idx - min_col
    field_offsets = [idx - min_col for idx in field_col_idxs]

    end_marker_n = norm(end_marker)
    results: list[dict[str, Any]] = []
    n_fields = len(column_fields)
    consecutive_sparse = 0

    for i, row_vals in enumerate(raw):
        row_num = start_row + i

        marker_val = norm(row_vals[entry_offset])
        if end_marker_n and end_marker_n in marker_val:
            break

        row_data: dict[str, Any] = {"_row": row_num}
        all_none = True
        for j, field_name in enumerate(field_names):
            val = row_vals[field_offsets[j]]
            row_data[field_name] = val
            if val is not None:
                all_none = False

        if all_none:
            consecutive_sparse += 1
            if consecutive_sparse >= SPARSE_ROW_BREAK_THRESHOLD:
                break
            continue

        if is_header_row(row_data):
            logger.debug("Skipping header row %d in sheet %r", row_num, ws.name)
            consecutive_sparse = 0
            continue

        non_row = [v for k, v in row_data.items()
                   if k != "_row" and v is not None]
        has_string = any(isinstance(v, str) for v in non_row)
        if not has_string and len(non_row) <= max(n_fields // 3, 1):
            consecutive_sparse += 1
            if consecutive_sparse >= SPARSE_ROW_BREAK_THRESHOLD:
                break
            continue

        consecutive_sparse = 0
        results.append(row_data)

    return results


# ---------------------------------------------------------------------------
# Strategy 3: Named range
# ---------------------------------------------------------------------------

def resolve_named_range(wb: xw.Book, name: str) -> Any:
    """Resolve a German Excel defined name to its value."""
    try:
        rng = wb.names[name].refers_to_range
        return rng.value
    except (KeyError, AttributeError):
        logger.warning("Named range %r not found in workbook", name)
        return None


# ---------------------------------------------------------------------------
# Strategy 4: Absolute address
# ---------------------------------------------------------------------------

def resolve_absolute(ws: xw.Sheet, address: str) -> Any:
    """Return the value at a fixed cell reference like 'C11'."""
    return ws.range(address).value


# ---------------------------------------------------------------------------
# Strategy 5: Column + row-offset within a block
# ---------------------------------------------------------------------------

def resolve_row_offset(
    ws: xw.Sheet,
    anchor_row: int,
    col: str,
    row_offset: int = 0,
) -> Any:
    """Return value at *col*, *anchor_row* + *row_offset*."""
    return cell_value(ws, col, anchor_row + row_offset)


# ---------------------------------------------------------------------------
# Strategy 6: Fixed result rows/cols
# ---------------------------------------------------------------------------

def resolve_fixed(ws: xw.Sheet, *, row: int, col: str) -> Any:
    """Read a fixed result location (typically formula outputs)."""
    return cell_value(ws, col, row)


# ---------------------------------------------------------------------------
# Item classifier
# ---------------------------------------------------------------------------

def classify_item(key: str, value: Any) -> str:
    """Classify an item as 'address', 'named_range', or 'config'."""
    if isinstance(value, str):
        if _is_cell_ref(value):
            return "address"
        if _is_named_range(value):
            return "named_range"
    return "config"
