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


def col_to_idx(col: str) -> int:
    """Convert column letters (A, ..., AA) to 1-based index."""
    result = 0
    for ch in col.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


def resolve_sheet_name(sheet_name: str, sheet_names: list[str]) -> str | None:
    """Case-insensitively resolve *sheet_name* against real workbook sheets.

    Excel doesn't allow two sheets to coexist with names differing only by
    case, so matching case-insensitively is always safe -- it can't
    introduce ambiguity between two distinct real sheets. Returns the
    actual, correctly-cased name (needed for exact-case downstream lookups
    like ``wb.sheets[name]``), or None if no sheet matches even
    case-insensitively.
    """
    target = sheet_name.casefold()
    for name in sheet_names:
        if name.casefold() == target:
            return name
    return None


def field_col(spec: str | dict) -> str:
    """Extract column letter from a field spec (string or dict with 'column')."""
    return spec if isinstance(spec, str) else spec.get("column", "A")


def is_formula(ws: xw.Sheet, col: str, row: int) -> bool:
    """Return True if a cell contains a formula."""
    f = ws.range((row, col_to_idx(col))).formula
    return isinstance(f, str) and f.startswith("=")


def cell_value(
    ws: xw.Sheet, col: str, row: int, *, skip_formulas: bool = False,
) -> Any:
    """Read a single cell by column letter(s) and row number.

    When *skip_formulas* is True, returns None for formula cells so
    only designer-entered input values are captured.
    """
    rng = ws.range((row, col_to_idx(col)))
    if skip_formulas:
        f = rng.formula
        if isinstance(f, str) and f.startswith("="):
            return None
    return rng.value


_BATCH_CHUNK_SIZE = 50


def _read_rect_chunked(
    ws: xw.Sheet, min_col: int, max_col: int, start_row: int, end_row: int,
    *, attr: str = "value",
) -> list[list[Any]]:
    """Read a rectangular region (all rows, min_col..max_col), row-chunked.

    xlwings' Mac AppleScript backend can silently drop a row's worth of
    values from a batch ``.value``/``.formula`` read when the range
    crosses certain hidden/grouped-row boundaries (confirmed empirically:
    requesting N rows sometimes returns N-1 values with no error),
    desynchronizing row-index arithmetic for every row after the drop.
    Reading in small chunks and validating each chunk's returned length
    against the expected row count catches this; any chunk that doesn't
    validate falls back to a per-row read for just that chunk, so the
    result stays correct without paying the per-row cost everywhere.
    """
    if start_row > end_row:
        return []
    rows: list[list[Any]] = []
    row = start_row
    while row <= end_row:
        chunk_end = min(row + _BATCH_CHUNK_SIZE - 1, end_row)
        expected_len = chunk_end - row + 1
        raw = getattr(ws.range((row, min_col), (chunk_end, max_col)), attr)
        if raw is None:
            # A fully empty/no-formula chunk -- preserve row count with
            # None-filled rows rather than dropping rows, so this chunk
            # still lines up 1:1 with a parallel value/formula read over
            # the same row range.
            raw = [[None] * (max_col - min_col + 1) for _ in range(expected_len)]
        elif not isinstance(raw, list):
            raw = [[raw]]
        elif not isinstance(raw[0], list):
            # xlwings returns a flat list for single-row OR single-column
            # ranges -- disambiguate before checking chunk length.
            raw = [[v] for v in raw] if min_col == max_col else [raw]

        if len(raw) != expected_len:
            logger.debug(
                "Batch %s read for cols %d-%d rows %d-%d returned %d rows, "
                "expected %d -- falling back to per-row read for this chunk",
                attr, min_col, max_col, row, chunk_end, len(raw), expected_len,
            )
            raw = []
            for r in range(row, chunk_end + 1):
                one = getattr(ws.range((r, min_col), (r, max_col)), attr)
                raw.append(one if isinstance(one, list) else [one])

        rows.extend(raw)
        row = chunk_end + 1
    return rows


def find_row_in_col(
    ws: xw.Sheet, col: str, needle: str, *,
    contains: bool = True, start_from: int = 1,
) -> int | None:
    """Return the first row where ``col``'s cell text matches *needle*.

    Reads the column in row-chunks for performance (see
    _read_rect_chunked for why a single unchunked batch read isn't safe).
    """
    needle_n = norm(needle)
    if not needle_n:
        return None
    col_idx = col_to_idx(col)
    last_row = ws.used_range.last_cell.row
    if start_from > last_row:
        return None
    rows = _read_rect_chunked(ws, col_idx, col_idx, start_from, last_row)
    for i, row_vals in enumerate(rows):
        cell_n = norm(row_vals[0])
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

def is_label_anchored_formula(
    ws: xw.Sheet, locator_col: str, locator_string: str, input_col: str,
    row_offset: int = 0,
) -> bool | None:
    """Return whether a label-anchored field's target cell is a formula.

    Returns None if the label can't be found (caller should skip the
    input/output cross-check rather than treat this as a mismatch).
    """
    row = find_row_in_col(ws, locator_col, locator_string)
    if row is None:
        return None
    return is_formula(ws, input_col, row + row_offset)


def resolve_label_anchored(
    ws: xw.Sheet,
    locator_col: str,
    locator_string: str,
    input_col: str,
    row_offset: int = 0,
    *,
    skip_formulas: bool = False,
) -> Any:
    """Find locator_string in locator_col, read input_col."""
    row = find_row_in_col(ws, locator_col, locator_string)
    if row is None:
        logger.warning(
            "Label %r not found in column %s of sheet %r",
            locator_string, locator_col, ws.name,
        )
        return None
    return cell_value(ws, input_col, row + row_offset,
                      skip_formulas=skip_formulas)


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
    skip_formulas: bool = False,
) -> list[dict[str, Any]]:
    """Iterate a repeating block, returning one dict per data row.

    Reads all needed columns in a single batch for performance.
    When *skip_formulas* is True, formula cells are returned as None.
    """
    entry_col = entry_locator.get("col") or header_locator.get("col") or "A"
    if not re.match(r"^[A-Za-z]{1,3}$", entry_col):
        # header_locator['col'] can hold a non-column placeholder (e.g. a
        # search string in the wrong field) when there's no real
        # entry_locator and the section is anchored purely by
        # entry_row_start -- fall back to a harmless default rather than
        # feeding a bogus column letter to col_to_idx().
        entry_col = "A"

    entry_string = entry_locator.get("string", "")

    def _discover_start_row() -> tuple[int | None, bool]:
        """Find the entry row by searching for the header + entry label.

        Returns (start_row, header_found) -- header_found distinguishes
        "header missing" from "entry label missing" for warning purposes.
        """
        if not entry_string:
            return None, False
        hdr_row = find_row_in_col(
            ws, header_locator["col"], header_locator["string"]
        )
        if hdr_row is None:
            return None, False
        start_row_found = find_row_in_col(
            ws, entry_col, entry_string, start_from=hdr_row,
        )
        if start_row_found is None:
            return None, True
        if is_entry_row_header(ws, start_row_found, column_fields):
            return max(start_row_found + 1, hdr_row), True
        return max(start_row_found, hdr_row), True

    if entry_row_start is not None:
        start_row = entry_row_start
        # entry_row_start always wins (it's the authoritative override), but
        # cross-check it against the discoverable label position -- if the
        # two disagree, that's a sign the hardcoded row has drifted from the
        # workbook's actual layout, so surface it instead of staying silent.
        if entry_string:
            discovered, _ = _discover_start_row()
            if discovered is not None and discovered != entry_row_start:
                logger.warning(
                    "entry_row_start=%d for entry label %r in sheet %r "
                    "disagrees with the discovered row %d -- using "
                    "entry_row_start, but the field map may be stale",
                    entry_row_start, entry_string, ws.name, discovered,
                )
    else:
        if not entry_string:
            return []
        discovered, header_found = _discover_start_row()
        if discovered is None:
            if not header_found:
                logger.warning(
                    "Block header %r not found in column %s",
                    header_locator["string"], header_locator["col"],
                )
            return []
        start_row = discovered

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

    # Read the rectangular region in row-chunks (see _read_rect_chunked --
    # a single unchunked batch read can silently drop rows across hidden
    # row boundaries, desynchronizing row-index arithmetic below).
    raw = _read_rect_chunked(ws, min_col, max_col, start_row, last_row)
    if not raw:
        return []

    # Optional: batch-read formulas to filter out formula cells
    formula_mask = None
    if skip_formulas:
        formula_mask = _read_rect_chunked(
            ws, min_col, max_col, start_row, last_row, attr="formula")

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

        # A bold entry-column cell is a section title (a totals row, or the
        # start of an unrelated table below that happens to reuse this
        # block's column layout), not more block data -- unlike
        # is_header_row() below, this doesn't depend on every mapped column
        # in the row happening to be a string, so it also catches rows that
        # mix a formula-cached number with text (e.g. a "Total ..." summary
        # row) and rows that otherwise look like valid data. Verified against
        # every header+entry block in both field-map versions: this is the
        # only bold row any block currently returns, so this can only ever
        # narrow (never break) existing results. This is a live per-cell
        # xlwings call (not part of the batch read above -- font.bold isn't
        # vectorizable across a range) but only fires for non-blank rows, and
        # measured at ~5ms/row even on the largest block (589 rows -> ~3s).
        if marker_val:
            entry_bold = ws.range((row_num, entry_col_idx)).font.bold
            if entry_bold:
                logger.debug(
                    "Stopping block scan at row %d in sheet %r -- entry "
                    "column is bold, signaling a new section", row_num,
                    ws.name,
                )
                break

        # Sparse/header detection always looks at raw (unfiltered) values,
        # so a row with real data isn't misclassified as blank just because
        # skip_formulas nulled out its formula-driven fields. row_data (the
        # returned dict) still applies the skip_formulas filter as before.
        row_data: dict[str, Any] = {"_row": row_num}
        raw_data: dict[str, Any] = {"_row": row_num}
        all_none = True
        for j, field_name in enumerate(field_names):
            raw_val = row_vals[field_offsets[j]]
            raw_data[field_name] = raw_val
            if raw_val is not None:
                all_none = False
            val = raw_val
            if formula_mask is not None:
                f = formula_mask[i][field_offsets[j]]
                if isinstance(f, str) and f.startswith("="):
                    val = None
            row_data[field_name] = val

        if all_none:
            consecutive_sparse += 1
            if consecutive_sparse >= SPARSE_ROW_BREAK_THRESHOLD:
                break
            continue

        if is_header_row(raw_data):
            logger.debug("Skipping header row %d in sheet %r", row_num, ws.name)
            consecutive_sparse = 0
            continue

        non_row = [v for k, v in raw_data.items()
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

_NOT_FOUND = object()


def _resolve_si_mirror_passthrough(wb: xw.Book, rng: xw.Range) -> Any:
    """Fall back to the base-tab counterpart of a "<Name> SI" formula cell.

    Some Excel-internal defined names (e.g. `Klima_Region`, `Klima_Standort`)
    resolve to a cell on the "<Name> SI" mirror tab whose formula (typically
    `=IF(ISTEXT(<Name>!<coord>),<Name>!<coord>,"")`) just passes through the
    real designer input that lives at the same coordinate on the base tab --
    this is Excel's own defined-name table pointing there, not a choice the
    field map's `sheet_name` makes, so the existing per-version sheet_name
    fix can't route around it. Only applies when the range's sheet actually
    ends in " SI" and the base tab's same-coordinate cell is itself not a
    formula; otherwise returns _NOT_FOUND so the caller keeps its existing
    behavior.
    """
    title = rng.sheet.name
    if not title.lower().endswith(" si"):
        return _NOT_FOUND
    base_title = resolve_sheet_name(
        title[: -len(" SI")], [s.name for s in wb.sheets])
    if base_title is None:
        return _NOT_FOUND
    base_rng = wb.sheets[base_title].range(rng.address)
    base_f = base_rng.formula
    if isinstance(base_f, str) and base_f.startswith("="):
        return _NOT_FOUND
    return base_rng.value


def resolve_named_range(
    wb: xw.Book, name: str, *, skip_formulas: bool = False,
) -> Any:
    """Resolve a German Excel defined name to its value."""
    try:
        rng = wb.names[name].refers_to_range
        if rng.shape != (1, 1):
            # Multi-cell destination -- PHPP defines several device-type-
            # name ranges broader than the single value they hold (e.g.
            # Kuehlgeraete_Kompressor_Umluft_Geraet -> K37:R37, every cell
            # blank except K37). Resolve to the top-left cell instead of
            # returning the whole nested-list .value/.formula, matching the
            # openpyxl-backend fix in PHX_pyxl.
            rng = rng[0, 0]
        if skip_formulas:
            f = rng.formula
            if isinstance(f, str) and f.startswith("="):
                base_value = _resolve_si_mirror_passthrough(wb, rng)
                if base_value is not _NOT_FOUND:
                    return base_value
                return None
        return rng.value
    except (KeyError, AttributeError):
        logger.warning("Named range %r not found in workbook", name)
        return None


# ---------------------------------------------------------------------------
# Strategy 4: Absolute address
# ---------------------------------------------------------------------------

def resolve_absolute(
    ws: xw.Sheet, address: str, *, skip_formulas: bool = False,
) -> Any:
    """Return the value at a fixed cell reference like 'C11'."""
    rng = ws.range(address)
    if skip_formulas:
        f = rng.formula
        if isinstance(f, str) and f.startswith("="):
            return None
    return rng.value


# ---------------------------------------------------------------------------
# Strategy 5: Column + row-offset within a block
# ---------------------------------------------------------------------------

def resolve_row_offset(
    ws: xw.Sheet,
    anchor_row: int,
    col: str,
    row_offset: int = 0,
    *,
    skip_formulas: bool = False,
) -> Any:
    """Return value at *col*, *anchor_row* + *row_offset*."""
    return cell_value(ws, col, anchor_row + row_offset,
                      skip_formulas=skip_formulas)


# ---------------------------------------------------------------------------
# Strategy 6: Fixed result rows/cols
# ---------------------------------------------------------------------------

def resolve_fixed(
    ws: xw.Sheet, *, row: int, col: str, skip_formulas: bool = False,
) -> Any:
    """Read a fixed result location (typically formula outputs)."""
    return cell_value(ws, col, row, skip_formulas=skip_formulas)
