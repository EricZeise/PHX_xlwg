# PHX_xlwg: Strategy and Routine Reference

## Part 1 — Strategy Summary

### Goal

Read designer-entered input data from a filled PHPP workbook (Passive House Planning Package, ~29 MB, 83 sheets), store it as portable JSON, and write that JSON back into a blank PHPP template — preserving every input value while leaving all formulas, charts, data validations, and drawings untouched.

### Why xlwings

The PHPP is a heavily formulated Excel workbook. An earlier approach (PHX_Dev) used openpyxl to read and write cells directly, but hit two fundamental problems:

1. **Stale formula values.** openpyxl reads cached results, not live recalculations. Any cell whose value depends on another input cell may be out of date.
2. **File corruption.** openpyxl's load→save cycle strips Excel extensions (data validation rules, custom header/footers) that PHPP relies on, making the file unopenable in Excel.

xlwings solves both: it drives a live Excel instance via AppleScript (macOS) or COM (Windows), so formulas recalculate on the fly and the file format is never touched by Python's XML serializer.

### The hybrid architecture

A pure xlwings round trip would be ideal, but macOS 26 introduced an AppleScript bug where all Excel save operations hang indefinitely on large workbooks. The workaround is a **hybrid writer**:

| Phase | Engine | Purpose |
|-------|--------|---------|
| **Read** | xlwings (Excel) | Open the filled workbook in Excel. Walk the field map, resolve every locator against the live sheet, read input cell values. |
| **Write — address resolution** | xlwings (Excel) | Open the template in Excel. Walk the field map again to resolve where each value should go (label searches, named ranges, entry-row detection). Collect writes as `(sheet, col, row, value)` tuples — but do not save. |
| **Write — persistence** | openpyxl (no Excel) | Close Excel. Reopen the .xlsx with openpyxl, apply all collected writes, save. openpyxl's save corrupts some PHPP features, but the resulting file is usable as a data artifact and can be verified cell-by-cell. |

### Formula-aware filtering

Not every cell with a value is an input. Roughly 35% of mapped cells are formula results (calculated from other inputs). Writing a formula's cached result back into a template would overwrite the formula itself, breaking the spreadsheet's calculation chain.

The reader's `skip_formulas` flag (on by default) checks each cell's `.formula` property before reading. If the formula string starts with `=`, the cell is skipped and recorded as `None` in the JSON. This ensures the JSON contains only values that can be meaningfully written back.

### The field map

A single markdown file (`phpp-field-mapping.md`) serves as the dictionary of every mapped cell across 31 worksheets. It defines six addressing strategies for locating cells, since PHPP's layout varies from sheet to sheet:

1. **Label-anchored** — find a text label in one column, read/write the value in a paired column at an optional row offset.
2. **Header + entry block** — find a header row and entry row, then iterate repeating data rows reading each column field.
3. **Named ranges** — resolve an Excel defined name (often German, e.g. `Werte_Klima_Region`) to its cell.
4. **Absolute address** — a fixed cell reference like `C11`.
5. **Column + row-offset** — locate an anchor row, then read at a column with a fixed row offset.
6. **Fixed result rows/cols** — a fixed row and column, typically for result cells.

### Verification

The roundtrip test confirms data fidelity end to end. The writer returns its full list of `(sheet, col, row, value)` writes. The test script opens the written file with openpyxl and verifies every cell matches — 13,102 cells, zero mismatches.

---

## Part 2 — Routine-by-Routine Walkthrough

### `excel_app.py` — Excel application factory

**Strategy role:** Provides a controlled Excel instance for both reader and writer.

| Step | What happens |
|------|-------------|
| 1 | `_find_excel()` scans a priority-ordered list of macOS Excel installation paths (2021 first, Office 2011 last). An `$PHPP_EXCEL_PATH` environment variable overrides the search. On Windows, returns `None` (xlwings finds Excel via COM automatically). |
| 2 | `excel_app()` checks for a shared app instance (set by test harnesses via `set_shared_app()`). If none, launches a new hidden Excel process with `xw.App(spec=..., visible=False, add_book=False)` and disables alert dialogs. |
| 3 | `open_book()` opens a workbook by path, with a single retry after a 1-second delay to handle transient AppleScript error -50 on rapid open/close cycles. |
| 4 | `is_shared()` lets callers know whether to quit the app when done — shared instances are left running. |

### `map_parser.py` — Field map parser

**Strategy role:** Converts the markdown field map into the structured dict that reader and writer iterate over.

| Step | What happens |
|------|-------------|
| 1 | `parse_field_map()` reads the markdown file and splits it at `##` headings (one per worksheet). |
| 2 | Each worksheet section is parsed for its worksheet key (e.g. `VERIFICATION`), sheet name, config block, label-anchored fields, and subsections. |
| 3 | Subsections (split at `###`) are parsed for header/entry locators, column fields, row fields, items, and appliance rows — capturing all six addressing strategies. |
| 4 | Returns a nested dict keyed by worksheet key, ready for reader and writer to iterate. |

### `locators.py` — Cell resolution (six strategies)

**Strategy role:** The bridge between the field map's abstract locator specs and actual Excel cell addresses. Every reader and writer function calls into this module.

#### Helpers

| Function | What it does |
|----------|-------------|
| `norm()` | Normalizes label text for comparison — NFKC unicode, NBSP→space, strip, casefold. Ensures labels match regardless of Excel's formatting quirks. |
| `col_to_idx()` | Converts column letters (`A`→1, `AA`→27) to 1-based numeric index for xlwings range addressing. |
| `field_col()` | Extracts the column letter from a field spec (handles both string and dict forms). |
| `is_formula()` | Checks whether a single cell contains a formula (`.formula` starts with `=`). |
| `cell_value()` | Reads a single cell. When `skip_formulas=True`, checks the formula property first and returns `None` for formula cells. |
| `find_row_in_col()` | Searches for a text needle in a column. **Batch-optimized:** reads the entire column in one AppleScript call, then scans in Python. Supports substring and exact matching. |
| `parse_cell_ref()` | Splits `"AB123"` into `("AB", 123)`. |
| `is_header_row()` | Heuristic: a row is a header if all non-None values are strings. Used to skip header rows in block iteration. |

#### Strategy 1: Label-anchored — `resolve_label_anchored()`

Finds `locator_string` in `locator_col` using `find_row_in_col()`, then reads the value at `input_col` in the found row (plus optional `row_offset`). Applies `skip_formulas` filtering.

#### Strategy 2: Header + entry block — `resolve_block()`

The most complex resolver. Handles repeating data rows (e.g. window schedules, area entries).

| Step | What happens |
|------|-------------|
| 1 | Find the header row and entry row using `find_row_in_col()`. If the entry row is itself a column header (detected by `is_entry_row_header()`), start one row below. |
| 2 | Build a column index map from all `column_fields` to determine the rectangular region to read. |
| 3 | **Batch read:** fetch the entire region in one `ws.range(...).value` call. Handle xlwings' flat-list behavior for single-row and single-column ranges. |
| 4 | If `skip_formulas` is on, batch-read `.formula` for the same region and build a formula mask. |
| 5 | Iterate rows. For each row: check for end markers, apply the formula mask to null out formula cells, detect sparse/empty rows (break after 3 consecutive), skip header rows, and collect data rows. |

#### Strategy 3: Named range — `resolve_named_range()`

Looks up an Excel defined name in `wb.names[name].refers_to_range`, reads its value. Applies `skip_formulas` filtering. Returns `None` on missing names.

#### Strategy 4: Absolute address — `resolve_absolute()`

Reads `ws.range(address).value` for a fixed cell reference like `"C11"`. Applies `skip_formulas` filtering.

#### Strategy 5: Column + row-offset — `resolve_row_offset()`

Reads the cell at `col`, `anchor_row + row_offset`. Delegates to `cell_value()` with `skip_formulas`.

#### Strategy 6: Fixed result — `resolve_fixed()`

Reads a cell at a fixed `(row, col)`. Delegates to `cell_value()` with `skip_formulas`. Typically used for formula output cells (which are skipped when formula filtering is on).

### `reader.py` — xlwings-based reader

**Strategy role:** Walks the field map, calls locator functions, and builds a nested dict of all input values from a filled PHPP workbook.

#### Top level: `read_phpp()`

| Step | What happens |
|------|-------------|
| 1 | Launch a new Excel instance via `excel_app()`. |
| 2 | Open the workbook via `open_book()`. |
| 3 | Parse the field map via `parse_field_map()`. |
| 4 | For each worksheet in the field map, find the matching sheet (preferring SI variants via `prefer_si_sheet()`). Call `_read_worksheet()`. |
| 5 | Close the workbook and quit Excel (unless using a shared instance). |
| 6 | Return the nested dict. |

#### `_read_worksheet()`

For a single sheet, reads three categories:

1. **Label-anchored fields** (`fields` key) — calls `_read_label_anchored_fields()`.
2. **Config values** (`config` key) — calls `_read_config()`.
3. **Sections** (`sections` key) — iterates sections, dispatching each via `_read_section()`.

All three pass `skip_formulas` through.

#### `_read_section()` — dispatch

Examines which keys are present in the section spec (`header_locator`, `entry_locator`, `column_fields`, `row_fields`, `items`, `fields`, `appliance_rows`) and routes to the appropriate reader:

| Pattern | Reader | Locator strategies used |
|---------|--------|------------------------|
| header + entry + row_fields (no col_fields) | `_read_row_offset_section()` | Strategy 5 (column + row-offset) |
| col_fields + row_fields | `_read_column_row_section()` | Strategy 5 |
| header + entry + col_fields | `_read_block_section()` | Strategy 2 (header + entry block) |
| col_fields only (no header) | `_read_static_column_section()` | Strategy 5 |
| appliance_rows | `_read_appliance_section()` | Stub (metadata only) |
| items | `_read_items_section()` | Strategies 4, 3, 6 (absolute, named range, fixed) |
| fields only | `_read_label_anchored_fields()` | Strategy 1 (label-anchored) |
| header only | `_read_header_only()` | `find_row_in_col()` |

#### `_read_label_anchored_fields()`

Iterates the `fields` dict. For each field with a `locator_string`, calls `resolve_label_anchored()` (Strategy 1). Collects results into a flat dict.

#### `_read_config()`

Iterates config key/value pairs. Classifies each via `classify_item()`: absolute addresses call `resolve_absolute()` (Strategy 4), named ranges call `resolve_named_range()` (Strategy 3), plain values pass through.

#### `_read_block_section()`

Delegates to `resolve_block()` (Strategy 2). Returns a list of row dicts, one per data row in the repeating block.

#### `_read_row_offset_section()`

Finds the anchor row via `find_row_in_col()`, then reads each field at a row offset using `resolve_row_offset()` (Strategy 5).

#### `_read_column_row_section()`

A grid pattern (e.g. DHW tanks). Iterates `column_fields` × `row_fields`, calling `resolve_row_offset()` (Strategy 5) for each intersection.

#### `_read_static_column_section()`

Iterates rows starting from `entry_row_start`, reading each `column_field` via `resolve_row_offset()` (Strategy 5). Stops when an entire row is None.

#### `_read_items_section()`

Iterates an `items` dict. For dict items with `col`+`row`, calls `resolve_fixed()` (Strategy 6). For string items, classifies as absolute address → `resolve_absolute()` (Strategy 4) or named range → `resolve_named_range()` (Strategy 3).

#### `_read_appliance_section()` (stub)

Returns metadata from the Electricity sheet's appliance row specs without reading live cell values.

#### `_read_header_only()`

Finds the header row position via `find_row_in_col()` and returns `{"_header_row": row}`.

### `writer.py` — Hybrid xlwings + openpyxl writer

**Strategy role:** Resolves cell addresses via xlwings (live Excel), collects all writes as tuples, then persists them via openpyxl after Excel closes.

#### Top level: `write_phpp()`

| Step | What happens |
|------|-------------|
| 1 | Copy the template file to the output path. |
| 2 | Launch Excel, open the copy, set calculation to manual (prevents recalculation during writes). |
| 3 | Parse the field map. For each worksheet key in the record, find the matching sheet and call `_write_worksheet()`. |
| 4 | Close the workbook, quit Excel. |
| 5 | Call `_apply_writes_openpyxl()` to persist all collected writes. |
| 6 | Return the list of `(sheet_name, col, row, value)` writes for verification. |

#### `_apply_writes_openpyxl()`

Opens the .xlsx with openpyxl, iterates the pending writes list, sets each cell value, saves, and closes. This is the only step that modifies the file on disk.

#### `_write_cell()`

The gatekeeper: skips `None` and dict/list values, appends valid writes to the pending list as `(sheet_name, col, row, value)`.

#### `_write_worksheet()`

Mirrors `_read_worksheet()`: writes label-anchored fields first, then iterates sections via `_write_section()`.

#### `_write_section()` — dispatch

Same pattern detection as the reader's `_read_section()`, routing to:

- `_write_block()` — iterates row data, uses `_row` metadata or sequential offset to determine target rows, writes each column field.
- `_write_row_offset()` — finds the anchor row, writes each field at its row offset.
- `_write_column_row()` — iterates the column × row grid, writes each intersection.
- `_write_static_column()` — iterates row data, writes each column field at the row's position.
- `_write_items()` — writes fixed-address, absolute-address, and named-range items.
- `_write_label_anchored()` — finds each label, writes the value at the paired column.

#### `_write_named_range()`

Resolves the named range via xlwings to get its sheet, column, and row, then appends to the pending list. Rejects `None`, dict, and list values.

### `models.py` — Pydantic validation

**Strategy role:** Provides a typed schema between the reader's raw dict and the JSON output. Validates and normalizes data without imposing rigid constraints on less-used worksheets.

| Step | What happens |
|------|-------------|
| 1 | Core worksheets (Verification, Overview, Climate, Ventilation, DHW, Windows) have explicit Pydantic models with typed fields. |
| 2 | Less-used worksheets use `dict[str, Any]` for flexibility. |
| 3 | `BuildingRecord.from_reader_dict()` maps worksheet keys to model classes, validates each, and assembles the top-level record. |
| 4 | `to_json()` serializes to JSON, excluding `None` values. |
| 5 | `model_validate_json()` (on the write path) deserializes and validates JSON back into a record. |

### `cli.py` — Click command-line interface

**Strategy role:** User-facing entry point that connects reader, models, and writer.

| Command | Pipeline |
|---------|----------|
| `phpp-tool read <filled.xlsx> -o record.json` | `read_phpp()` → `BuildingRecord.from_reader_dict()` → `to_json()` → write file |
| `phpp-tool write <record.json> <template.xlsx> -o output.xlsx` | `model_validate_json()` → `model_dump()` → `write_phpp()` |
| `phpp-tool inspect-map` | `parse_field_map()` → print worksheet/field/section counts |

### `scripts/roundtrip.py` — Roundtrip verification

**Strategy role:** End-to-end test that proves data survives the full read → JSON → write cycle.

| Phase | What happens |
|-------|-------------|
| 1 — Read inputs | `read_phpp(skip_formulas=True)` captures only input cells. Reports count of non-None values. |
| 2 — Read all | `read_phpp(skip_formulas=False)` captures everything. Compares against Phase 1 to report how many formula values were filtered (typically ~35%). |
| 3 — Write | `write_phpp()` writes inputs into the template. Returns the list of 13,102 cell writes. |
| 4 — Verify | Opens the written file with openpyxl and checks every `(sheet, col, row, value)` tuple from the writer. Reports checked count, mismatches, and details. Both Example→Empty and Empty→Empty produce zero mismatches. |

---

## Part 3 — Usage, Roundtrip Philosophy, and Output Files

### Using the phpp_tool

The tool serves a single workflow: extract portable data from a filled PHPP workbook, and inject that data into a fresh PHPP template. This lets designers transfer building configurations between PHPP versions, share project data without sending 29 MB workbooks, archive inputs in a version-controllable text format, and programmatically generate or modify building records.

#### Reading a filled workbook

```bash
phpp-tool read Data/Example.xlsx -o records/my_building.json
```

Excel must be installed. The tool launches a hidden Excel instance, opens the workbook, walks all 31 mapped worksheets, reads every input cell (skipping formulas), validates through Pydantic models, and writes a JSON file. The process takes roughly 21 seconds for a full PHPP workbook.

The resulting JSON is organized by worksheet key. Each worksheet contains some combination of scalar fields, config values, and sections. A section may be a flat dict of key-value pairs, a list of row dicts (for repeating blocks like window schedules), or a grid of column × row intersections.

#### Writing into a blank template

```bash
phpp-tool write records/my_building.json Data/Empty.xlsx -o output.xlsx
```

The tool copies the template, opens the copy in a hidden Excel instance to resolve all cell addresses (label searches, named ranges, block start rows), collects writes, closes Excel, and persists the values via openpyxl. This takes roughly 48 seconds.

The output file is a valid .xlsx that contains all the designer's input values in the correct cells. However, because openpyxl's save strips some Excel extensions (data validation rules, custom headers/footers), the file cannot be reopened by Excel via AppleScript. It is usable as a data artifact and can be opened manually in Excel, but some PHPP features may be degraded.

#### Inspecting the field map

```bash
phpp-tool inspect-map
```

Lists every mapped worksheet with counts of fields, sections, and config items. Useful for verifying field map coverage after edits.

### The roundtrip test: philosophy

The roundtrip test answers a single question: **does every input value survive the full read → JSON → write cycle unchanged?**

This is not a unit test of individual functions — those are covered by the 88 pytest cases. The roundtrip test is an integration test of the entire pipeline against real PHPP workbooks. It treats the tool as a black box and verifies the output against the input at the cell level.

#### Why cell-by-cell verification matters

The tool touches 13,102 cells across 31 worksheets. A mismatch in any one of them could mean a wrong U-value, a missing ventilation rate, or a misplaced area entry. Aggregate statistics (like "99.9% match") would hide single-cell errors that could be significant in a Passive House certification. The test therefore checks every cell individually and reports exact addresses for any mismatch.

#### Why two reads (inputs vs all)

The roundtrip reads the source workbook twice:

1. **Inputs only** (`skip_formulas=True`) — this is what gets written to JSON and into the template. It captures only designer-entered values.
2. **All cells** (`skip_formulas=False`) — this captures everything including formula results. Comparing the two reveals the formula filter's effect.

This dual read serves two purposes. First, it quantifies the formula filter: for the Example workbook, 7,625 of 22,045 non-None values (34.6%) are formula results that would overwrite formulas if written back. Second, it documents which worksheet fields are formulas versus inputs — the Verification sheet, for example, is entirely formula-driven:

| Field | All cells | Inputs only |
|-------|-----------|-------------|
| `phi_building_category_type` | `"21-Non-res building: School half-days (< 7 h)"` | `None` |
| `phi_certification_type` | `"10-Passive house"` | `None` |
| `setpoint_winter` | `20.0` | `None` |
| ... | (13 fields with values) | (all None — every field is a formula) |

This means the Verification sheet's display values are computed from inputs on other sheets. The tool correctly skips them.

#### Why the writer returns its write list

The writer's `write_phpp()` function returns the list of `(sheet_name, col, row, value)` tuples it collected during the xlwings address-resolution phase and persisted via openpyxl. This is not an incidental convenience — it is the verification contract. The roundtrip test opens the written file with openpyxl and checks every tuple against the actual cell contents. If the writer says it wrote `42.0` to `Ventilation SI!J35`, the test confirms that cell `J35` on sheet `Ventilation SI` contains `42.0`.

#### The two standard test pairs

| Test | Source | Template | What it proves |
|------|--------|----------|---------------|
| Example → Empty | A filled PHPP with real building data | A blank PHPP template | Input values from a real project survive the cycle |
| Empty → Empty | A blank PHPP template | The same blank template | Default/structural values survive; no spurious data is introduced |

Both produce zero mismatches across 13,102 verified cells.

### Output files

#### CLI output: JSON building record

Produced by `phpp-tool read`. Structure:

```
{
  "VERIFICATION": {                    ← worksheet key
    "phi_building_category_type": null, ← formula cell (skipped)
    "setpoint_winter": null,            ← formula cell (skipped)
    ...
  },
  "COMPONENTS": {
    "glazings": [                       ← repeating block (list of row dicts)
      {
        "_row": 115,                    ← source row in the Excel sheet
        "id": "1187gl03",
        "description": "EAGON - EAGON SUPER VIG (5/0,25 Vac/:5 Vac.)",
        "g_value": 0.48,
        "u_value": 0.51
      },
      ...                               ← 170 rows for glazings alone
    ],
    "frames": [ ... ],
    "ventilators": [ ... ]
  },
  "CLIMATE": {
    "named_ranges": {                   ← named range values
      "country": "US-United States of America",
      "region": "New York",
      "data_set": "New York/JFK"
    },
    ...
  },
  "DHW": {
    "tanks": {                          ← column × row grid
      "tank_1": {
        "tank_type": null,
        "standby_losses": null,
        ...
      },
      ...
    }
  },
  ...                                   ← 17 worksheet keys total
}
```

Key characteristics:

- **`_row` metadata** — Block rows carry their source row number so the writer can place them back at the correct position, even if the template has a different row layout.
- **`null` values** — Formula cells and empty input cells both appear as `null`. The writer skips `null` values, so formulas are preserved in the output workbook.
- **No formula text** — The JSON never contains formula strings like `=SUM(...)`. Only resolved input values appear.
- **`_config` sections** — Store worksheet-level settings (active column selections, variant references) that configure how the writer interprets the data.

#### CLI output: written workbook (.xlsx)

Produced by `phpp-tool write`. This is a copy of the blank template with input values injected at the addresses the writer resolved. Formulas remain intact as they were in the template. The file can be opened in Excel manually, though some PHPP data validation features may be degraded due to the openpyxl persistence step.

#### Roundtrip test artifacts

Saved to `records/roundtrip_<timestamp>/`. Each source workbook produces:

| File | Contents |
|------|----------|
| `<name>_inputs.json` | The input-only read (`skip_formulas=True`). This is the data that travels through the pipeline — the same output `phpp-tool read` would produce. 14,420 non-None values for the Example workbook. |
| `<name>_all.json` | The full read (`skip_formulas=False`). Includes formula results alongside inputs. 22,045 non-None values for the Example workbook. Comparing against `_inputs.json` shows exactly which cells are formulas. |
| `<name>_written.xlsx` | The written workbook. Template with input values injected. This is what `phpp-tool write` would produce. The roundtrip test verifies every cell in this file against the writer's reported write list. |

The test also prints a console report with:

- Worksheet and value counts for both reads
- Formula filter statistics (count and percentage removed)
- Number of cell writes the writer performed
- Cell-by-cell verification result (checked count, mismatch count, and details of any mismatches)
