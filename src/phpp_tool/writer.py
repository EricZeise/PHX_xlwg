"""Write a building record into a PHPP workbook via xlwings + surgical XML.

xlwings opens the workbook in a live Excel instance to resolve locator
lookups (label searches, named ranges, entry-row detection). Cell values
are NOT saved via Excel — on macOS 26, all AppleScript save paths hang
on large workbooks. Writes are collected during the xlwings pass, Excel
closes without saving (so the on-disk copy stays byte-identical to the
template), and then `surgical_writer.apply_surgical_writes()` persists the
writes by editing the .xlsx as a ZIP archive rather than going through
openpyxl's load->save cycle. This preserves <extLst> extensions (e.g. Data
Validation) and <headerFooter> content that an openpyxl save would drop.
"""

from __future__ import annotations

import logging
import shutil
from pathlib import Path
from typing import Any

import xlwings as xw

from phpp_tool.excel_app import excel_app, is_shared, open_book
from phpp_tool.locators import (
    field_col,
    find_row_in_col,
    is_entry_row_header,
    parse_cell_ref,
)
from phpp_tool.map_parser import parse_field_map
from phpp_tool.surgical_writer import apply_surgical_writes

logger = logging.getLogger(__name__)


def write_phpp(
    record: dict[str, Any],
    template_path: str | Path,
    output_path: str | Path,
    field_map_path: str | Path = "phpp-field-mapping.md",
) -> list[tuple[str, str, int, Any]]:
    """Write a building record into a PHPP workbook template.

    Returns the list of writes performed as (sheet_name, col, row, value)
    tuples, for verification by callers.
    """
    template_path = Path(template_path)
    output_path = Path(output_path)

    shutil.copy2(template_path, output_path)
    output_path.chmod(0o644)
    pending: list[tuple[str, str, int, Any]] = []

    app = excel_app()
    try:
        wb = open_book(app, str(output_path.resolve()))
        app.calculation = "manual"

        field_map = parse_field_map(field_map_path)
        sheet_names = [s.name for s in wb.sheets]
        total_writes = 0

        for ws_key, ws_data in record.items():
            if ws_data is None:
                continue
            ws_spec = field_map.get(ws_key)
            if ws_spec is None:
                logger.info("No field map entry for %r, skipping", ws_key)
                continue
            sheet_name = ws_spec["sheet_name"]
            if sheet_name not in sheet_names:
                logger.warning("Sheet %r not in template, skipping %s",
                               sheet_name, ws_key)
                continue
            ws = wb.sheets[sheet_name]
            total_writes += _write_worksheet(ws, wb, ws_spec, ws_data, pending)

        wb.close()
    finally:
        if not is_shared(app):
            app.quit()

    apply_surgical_writes(template_path, output_path, pending)
    logger.info("Wrote %d cell values", total_writes)
    return pending


# ======================================================================
# Cell writing
# ======================================================================

def _write_cell(
    ws: xw.Sheet, col: str, row: int, value: Any,
    pending: list[tuple[str, str, int, Any]],
) -> bool:
    """Record a cell write for later openpyxl persistence. Returns True if recorded."""
    if value is None:
        return False
    if isinstance(value, (dict, list)):
        return False
    pending.append((ws.name, col, row, value))
    return True


# ======================================================================
# Worksheet dispatch
# ======================================================================

def _write_worksheet(
    ws: xw.Sheet,
    wb: xw.Book,
    ws_spec: dict[str, Any],
    ws_data: dict[str, Any],
    pending: list[tuple[str, str, int, Any]],
) -> int:
    """Write all mapped values into a single worksheet. Returns write count."""
    count = 0
    if ws_spec.get("fields"):
        count += _write_label_anchored(ws, ws_spec["fields"], ws_data, pending)

    for sec_name, sec_spec in ws_spec.get("sections", {}).items():
        sec_data = ws_data.get(sec_name)
        if sec_data is None:
            continue
        count += _write_section(ws, wb, sec_spec, sec_data, pending)
    return count


# --- Strategy 1: label-anchored fields ---

def _write_label_anchored(
    ws: xw.Sheet,
    fields: dict[str, dict],
    data: dict[str, Any],
    pending: list[tuple[str, str, int, Any]],
) -> int:
    count = 0
    for field_name, spec in fields.items():
        value = data.get(field_name)
        if value is None or not spec.get("locator_string"):
            continue
        row = find_row_in_col(ws, spec["locator_col"], spec["locator_string"])
        if row is None:
            logger.warning("Label %r not found for write",
                           spec["locator_string"])
            continue
        if _write_cell(ws, spec["input_col"],
                       row + spec.get("row_offset", 0), value, pending):
            count += 1
    return count


# --- Section dispatch ---

def _write_section(
    ws: xw.Sheet,
    wb: xw.Book,
    sec_spec: dict[str, Any],
    sec_data: Any,
    pending: list[tuple[str, str, int, Any]],
) -> int:
    has_header = "header_locator" in sec_spec
    has_entry = "entry_locator" in sec_spec
    has_col_fields = "column_fields" in sec_spec
    has_row_fields = "row_fields" in sec_spec
    has_items = "items" in sec_spec
    has_fields = "fields" in sec_spec

    items = sec_spec.get("items", {})
    entry_row_start = (items.get("entry_row_start")
                       or items.get("entry_start_row")
                       or items.get("start_row"))

    if has_header and has_entry and has_row_fields and not has_col_fields:
        return _write_row_offset(ws, sec_spec, sec_data, pending)
    if has_col_fields and has_row_fields:
        return _write_column_row(ws, sec_spec, sec_data, entry_row_start,
                                 pending)
    if has_header and has_entry and has_col_fields:
        return _write_block(ws, sec_spec, sec_data, entry_row_start, pending)
    if has_col_fields and not has_header:
        return _write_static_column(ws, sec_spec, sec_data, entry_row_start,
                                    pending)
    if has_col_fields and entry_row_start is not None:
        return _write_block(ws, sec_spec, sec_data, entry_row_start, pending)
    if has_items:
        return _write_items(ws, wb, items, sec_spec.get("items_kind", {}),
                            sec_data, pending)
    if has_fields:
        return _write_label_anchored(ws, sec_spec["fields"], sec_data,
                                     pending)
    return 0


# --- Strategy 2: repeating block ---

def _write_block(
    ws: xw.Sheet,
    sec_spec: dict[str, Any],
    rows_data: list[dict[str, Any]],
    entry_row_start: int | None,
    pending: list[tuple[str, str, int, Any]],
) -> int:
    if not isinstance(rows_data, list):
        return 0

    if entry_row_start is None:
        entry_locator = sec_spec.get("entry_locator", {})
        entry_string = entry_locator.get("string", "")
        if entry_string:
            entry_col = entry_locator.get(
                "col", sec_spec.get("header_locator", {}).get("col", "A"))
            hdr = sec_spec.get("header_locator", {})
            hdr_row = 1
            if hdr.get("string"):
                hdr_found = find_row_in_col(
                    ws, hdr.get("col", "A"), hdr["string"])
                if hdr_found is not None:
                    hdr_row = hdr_found
            found = find_row_in_col(
                ws, entry_col, entry_string, start_from=hdr_row)
            if found is not None:
                col_fields = sec_spec.get("column_fields", {})
                if is_entry_row_header(ws, found, col_fields):
                    entry_row_start = found + 1
                else:
                    entry_row_start = found

    has_row_metadata = all(
        "_row" in rd for rd in rows_data) if rows_data else False
    if entry_row_start is None and not has_row_metadata:
        logger.warning("Cannot determine block start row for write")
        return 0

    count = 0
    column_fields = sec_spec["column_fields"]
    for i, row_data in enumerate(rows_data):
        if "_row" in row_data:
            target_row = row_data["_row"]
        elif entry_row_start is not None:
            target_row = entry_row_start + i
        else:
            continue
        for field_name, field_spec in column_fields.items():
            if _write_cell(ws, field_col(field_spec), target_row,
                           row_data.get(field_name), pending):
                count += 1
    return count


# --- Row-offset section (DHW recirc/branch piping) ---

def _write_row_offset(
    ws: xw.Sheet,
    sec_spec: dict[str, Any],
    data: dict[str, Any],
    pending: list[tuple[str, str, int, Any]],
) -> int:
    if not isinstance(data, dict):
        return 0
    entry_loc = sec_spec["entry_locator"]
    anchor_row = find_row_in_col(ws, entry_loc["col"], entry_loc["string"])
    if anchor_row is None:
        logger.warning("Entry locator %r not found for write",
                       entry_loc["string"])
        return 0
    input_col = sec_spec.get("items", {}).get("input_col_start", "J")
    count = 0
    for field_name, field_spec in sec_spec["row_fields"].items():
        value = data.get(field_name)
        if value is None:
            continue
        offset = field_spec.get("row_offset", field_spec.get("row", 0))
        if _write_cell(ws, input_col, anchor_row + offset, value, pending):
            count += 1
    return count


# --- Column + row section (DHW tanks) ---

def _write_column_row(
    ws: xw.Sheet,
    sec_spec: dict[str, Any],
    data: dict[str, Any],
    entry_row_start: int | None,
    pending: list[tuple[str, str, int, Any]],
) -> int:
    if not isinstance(data, dict):
        return 0
    if entry_row_start is None:
        entry_loc = sec_spec.get("entry_locator", {})
        if entry_loc:
            entry_row_start = find_row_in_col(
                ws, entry_loc["col"], entry_loc["string"])
        if entry_row_start is None:
            return 0

    count = 0
    for col_name, col_spec in sec_spec["column_fields"].items():
        entity_data = data.get(col_name)
        if not isinstance(entity_data, dict):
            continue
        col_letter = field_col(col_spec)
        for field_name, field_spec in sec_spec["row_fields"].items():
            value = entity_data.get(field_name)
            if value is None:
                continue
            offset = field_spec.get("row_offset", field_spec.get("row", 0))
            if _write_cell(ws, col_letter, entry_row_start + offset, value,
                           pending):
                count += 1
    return count


# --- Static column section ---

def _write_static_column(
    ws: xw.Sheet,
    sec_spec: dict[str, Any],
    data: list[dict[str, Any]] | dict[str, Any],
    entry_row_start: int | None,
    pending: list[tuple[str, str, int, Any]],
) -> int:
    if not isinstance(data, list):
        return 0
    items = sec_spec.get("items", {})
    start_row = entry_row_start or items.get("entry_start_row")
    if start_row is None:
        return 0
    count = 0
    column_fields = sec_spec["column_fields"]
    for i, row_data in enumerate(data):
        target_row = row_data.get("_row", start_row + i)
        for field_name, field_spec in column_fields.items():
            if _write_cell(ws, field_col(field_spec), target_row,
                           row_data.get(field_name), pending):
                count += 1
    return count


# --- Items section (Overview, SolarDHW, CoolingUnits, PER) ---

def _write_items(
    ws: xw.Sheet,
    wb: xw.Book,
    items_spec: dict[str, Any],
    items_kind: dict[str, str],
    data: dict[str, Any],
    pending: list[tuple[str, str, int, Any]],
) -> int:
    if not isinstance(data, dict):
        return 0
    count = 0
    for key, spec_value in items_spec.items():
        value = data.get(key)
        if value is None:
            continue
        if isinstance(spec_value, dict):
            if "col" in spec_value and "row" in spec_value:
                if _write_cell(ws, spec_value["col"],
                               spec_value["row"], value, pending):
                    count += 1
            continue
        kind = items_kind.get(key, "literal")
        if kind == "address":
            col, row = parse_cell_ref(spec_value)
            if _write_cell(ws, col, row, value, pending):
                count += 1
        elif kind == "named_range":
            if _write_named_range(wb, spec_value, value, pending):
                count += 1
    return count


def _write_named_range(
    wb: xw.Book, name: str, value: Any,
    pending: list[tuple[str, str, int, Any]],
) -> bool:
    """Resolve a named range via xlwings and record the write."""
    if value is None or isinstance(value, (dict, list)):
        return False
    try:
        rng = wb.names[name].refers_to_range
        from openpyxl.utils import get_column_letter
        col_letter = get_column_letter(rng.column)
        pending.append((rng.sheet.name, col_letter, rng.row, value))
        return True
    except (KeyError, AttributeError):
        logger.warning("Named range %r not found for write", name)
        return False
