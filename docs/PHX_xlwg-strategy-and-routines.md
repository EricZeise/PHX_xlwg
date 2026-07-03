# PHX_xlwg: Strategy and Routine Reference

## Part 1 — Strategy Summary

### Goal

Read designer-entered input data from a filled PHPP workbook (Passive House Planning Package, ~29 MB, 83 sheets), store it as portable JSON, and write that JSON back into a blank PHPP template — preserving every input value while leaving all formulas, charts, data validations, and drawings untouched.

### Why xlwings

The PHPP is a heavily formulated Excel workbook. An earlier approach (PHX_Dev) used openpyxl to read cells directly, which has one fundamental problem xlwings solves: **stale formula values** — openpyxl reads cached results, not live recalculations, so any cell whose value depends on another input cell may be out of date. xlwings solves this by driving a live Excel instance via AppleScript (macOS) or COM (Windows), so formulas recalculate on the fly.

(PHX_Dev's *write* side had a related historical problem — an early naive openpyxl load→save cycle stripped Excel extensions, data validation rules, and custom headers/footers — but PHX_Dev resolved this itself with a surgical ZIP/XML patch that edits only the `<sheetData>` region of affected sheets, never touching `<extLst>`/`<headerFooter>` content. Both PHX_pyxl and PHX_xlwg have since adopted the same mechanism — see `surgical_writer.py` in the routine walkthrough below — so file corruption on write is not a reason to prefer xlwings over openpyxl; only live formula recalculation is.)

### The hybrid architecture

A pure xlwings round trip would be ideal, but macOS 26 introduced an AppleScript bug where all Excel save operations hang indefinitely on large workbooks. The workaround is a **hybrid writer**:

| Phase | Engine | Purpose |
|-------|--------|---------|
| **Read** | xlwings (Excel) | Open the filled workbook in Excel. Walk the field map, resolve every locator against the live sheet, read input cell values. |
| **Write — address resolution** | xlwings (Excel) | Open the template in Excel. Walk the field map again to resolve where each value should go (label searches, named ranges, entry-row detection). Collect writes as `(sheet, col, row, value)` tuples — but do not save. |
| **Write — persistence** | `surgical_writer.py` (no Excel) | Close Excel (without saving — the on-disk copy stays byte-identical to the template). Patch the `.xlsx` as a ZIP archive: edit only the `<sheetData>` region of sheets with writes via `lxml`, leaving `<extLst>`/`<headerFooter>`/everything else as untouched original bytes. Verified byte-for-byte across all 83 sheets of a full roundtrip write. |

### Formula-aware filtering

Not every cell with a value is an input. Roughly 35% of mapped cells are formula results (calculated from other inputs). Writing a formula's cached result back into a template would overwrite the formula itself, breaking the spreadsheet's calculation chain.

The reader's `skip_formulas` flag (on by default) checks each cell's `.formula` property before reading. If the formula string starts with `=`, the cell is skipped and recorded as `None` in the JSON. This ensures the JSON contains only values that can be meaningfully written back.

### The field map

A markdown file per PHPP workbook type — `phpp-field-mapping/EN_10_6_IP.md` and `EN_10_6_SI.md`, selected via `--phpp-version` — serves as the dictionary of every mapped cell across 31 worksheets. It defines six addressing strategies for locating cells, since PHPP's layout varies from sheet to sheet:

1. **Label-anchored** — find a text label in one column, read/write the value in a paired column at an optional row offset.
2. **Header + entry block** — find a header row and entry row, then iterate repeating data rows reading each column field.
3. **Named ranges** — resolve an Excel defined name (often German, e.g. `Werte_Klima_Region`) to its cell.
4. **Absolute address** — a fixed cell reference like `C11`.
5. **Column + row-offset** — locate an anchor row, then read at a column with a fixed row offset.
6. **Fixed result rows/cols** — a fixed row and column, typically for result cells.

### Verification

The roundtrip test confirms data fidelity end to end. The writer returns its full list of `(sheet, col, row, value)` writes. The test script opens the written file with openpyxl and verifies every cell matches — 13,102 cells, zero mismatches.

---

## Part 2 — Using the Routines and Scripts

This section shows exactly how to invoke everything described in Part 1: the `phpp-tool` CLI and the roundtrip script. All commands assume Excel is installed, an activated venv, and a working directory at the project root.

```bash
cd /Users/smini/Documents/Coding/PHX_xlwg
source .venv/bin/activate
```

**If `.venv/` is missing or broken** (e.g. after a folder rename — see the known venv-path issue), recreate it before running anything:

```bash
rm -rf .venv
python3.13 -m venv .venv
.venv/bin/pip install -e ".[dev]"
```

### `phpp-tool read` — extract a filled workbook to JSON

```bash
phpp-tool read WORKBOOK -o OUTPUT [--phpp-version PHPP_VERSION] [--field-map FIELD_MAP]
```

- `WORKBOOK` (positional, required) — path to the filled `.xlsx` to read.
- `-o, --output OUTPUT` — JSON output path (omit to print to stdout).
- `--phpp-version PHPP_VERSION` — resolves `phpp-field-mapping/<version>.md` (default `EN_10_6_IP`). Use `EN_10_6_SI` for a genuinely SI-native single-shell workbook.
- `--field-map FIELD_MAP` — direct-path override, bypassing `--phpp-version` entirely.

Example:

```bash
phpp-tool read Data/Example_IP.xlsx -o records/my_building.json
```

Internally: launches a hidden Excel instance → `read_phpp()` → `BuildingRecord.from_reader_dict()` → `to_json()`, then stamps the output JSON with `_phpp_version`. Takes roughly 21 seconds for a full PHPP workbook, all of it live in Excel — no cached-value staleness.

### `phpp-tool write` — inject a JSON record into a blank template

```bash
phpp-tool write RECORD_FILE TEMPLATE -o OUTPUT [--phpp-version PHPP_VERSION] [--field-map FIELD_MAP]
```

- `RECORD_FILE` (positional, required) — the JSON produced by `read`.
- `TEMPLATE` (positional, required) — the blank `.xlsx` to write into.
- `-o, --output OUTPUT` (required) — path for the written workbook.
- `--phpp-version PHPP_VERSION` / `--field-map FIELD_MAP` — same as `read`. **Must match the version the record was read with** — `write` compares `--phpp-version` against the record's stamped `_phpp_version` and prints a warning (not a hard error) on mismatch.

Example:

```bash
phpp-tool write records/my_building.json Data/Empty_IP.xlsx -o output.xlsx
```

Internally: `model_validate_json()` → `model_dump(exclude_none=True)` → `write_phpp()`, which resolves addresses live in Excel, then persists values via `surgical_writer.py`'s ZIP/XML patch after Excel closes without saving (the macOS 26 AppleScript-save workaround, and the mechanism that preserves `<extLst>`/`<headerFooter>` content). Takes roughly 48 seconds.

### `phpp-tool inspect-map` — audit field map coverage

```bash
phpp-tool inspect-map [--phpp-version PHPP_VERSION] [--field-map FIELD_MAP]
```

- `--phpp-version PHPP_VERSION` — resolves `phpp-field-mapping/<version>.md` (default `EN_10_6_IP`).
- `--field-map FIELD_MAP` — direct-path override, bypassing `--phpp-version` entirely.

Example:

```bash
phpp-tool inspect-map --phpp-version EN_10_6_IP
```

Prints every mapped worksheet with its field/section/config counts. Use this after editing a field map file to confirm the parser still finds everything (and that every config/items entry still has a valid type tag — `inspect-map` will raise `FieldMapError` immediately if one is missing).

### `scripts/roundtrip.py` — end-to-end verification

```bash
python scripts/roundtrip.py [--phpp-version PHPP_VERSION] SOURCE TEMPLATE [SOURCE TEMPLATE ...]
```

- `--phpp-version PHPP_VERSION` — resolves `phpp-field-mapping/<version>.md` (default `EN_10_6_IP`); if given, must appear before the `SOURCE`/`TEMPLATE` pairs.
- `SOURCE` (positional, required, repeatable) — filled `.xlsx` to read from.
- `TEMPLATE` (positional, required, repeatable) — blank `.xlsx` to write into, paired with the `SOURCE` immediately before it. Pass additional `SOURCE TEMPLATE` pairs on the same command line to run several roundtrips in one invocation (results print a combined summary at the end).

Example:

```bash
python scripts/roundtrip.py Data/Example_IP.xlsx Data/Empty_IP.xlsx
python scripts/roundtrip.py Data/Empty_IP.xlsx Data/Empty_IP.xlsx
```

Output artifacts land in `records/roundtrip_<timestamp>/`. Requires Excel throughout (read, address resolution, and the final openpyxl-based cell verification).

| Phase | What runs |
|-------|-----------|
| 1 — Read inputs | `read_phpp(skip_formulas=True)` via xlwings |
| 2 — Read all | `read_phpp(skip_formulas=False)` via xlwings — reports formula filter stats |
| 3 — Write | xlwings address resolution + surgical XML persistence (`surgical_writer.py`) |
| 4 — Verify | openpyxl reads the written file and checks every writer-reported cell |

Unlike PHX_pyxl's roundtrip script, there is no separate "Part 2" Excel-cache-validation phase here — every read in this project already goes through live Excel, so there's no cache to validate against.

### Running the test suite

```bash
pytest tests/ -v          # requires Excel, ~9s
```

### Quick reference

| Task | Command |
|------|---------|
| Read a filled PHPP → JSON | `phpp-tool read Data/Example_IP.xlsx -o records/my_building.json` |
| Write JSON → blank PHPP | `phpp-tool write records/my_building.json Data/Empty_IP.xlsx -o output.xlsx` |
| Check field map coverage | `phpp-tool inspect-map` |
| Roundtrip verification | `python scripts/roundtrip.py Data/Example_IP.xlsx Data/Empty_IP.xlsx` |
| Unit tests | `pytest tests/ -v` |

---

## Part 3 — Routine-by-Routine Walkthrough

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

### `writer.py` — Hybrid xlwings address resolution + surgical XML writer

**Strategy role:** Resolves cell addresses via xlwings (live Excel), collects all writes as tuples, then hands them to `surgical_writer.apply_surgical_writes()` for persistence after Excel closes.

#### Top level: `write_phpp()`

| Step | What happens |
|------|-------------|
| 1 | Copy the template file to the output path. |
| 2 | Launch Excel, open the copy, set calculation to manual (prevents recalculation during writes). |
| 3 | Parse the field map. For each worksheet key in the record, find the matching sheet and call `_write_worksheet()`. |
| 4 | Close the workbook **without saving** (xlwings' `Book.close()` never saves), quit Excel — the on-disk copy remains byte-identical to the template. |
| 5 | Call `surgical_writer.apply_surgical_writes(template_path, output_path, pending)` to persist all collected writes via ZIP/XML patch. |
| 6 | Return the list of `(sheet_name, col, row, value)` writes for verification. |

#### `surgical_writer.py`

Same module as PHX_pyxl's — see its routine walkthrough there. Patches only the `<sheetData>` region of sheets with writes, via `lxml`, leaving `<extLst>`/`<headerFooter>`/everything else in the ZIP archive untouched. This is what actually modifies the file on disk; xlwings never saves.

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

## Part 4 — Usage, Roundtrip Philosophy, and Output Files

### Using the phpp_tool

The tool serves a single workflow: extract portable data from a filled PHPP workbook, and inject that data into a fresh PHPP template. This lets designers transfer building configurations between PHPP versions, share project data without sending 29 MB workbooks, archive inputs in a version-controllable text format, and programmatically generate or modify building records.

#### Reading a filled workbook

```bash
phpp-tool read Data/Example_IP.xlsx -o records/my_building.json
```

Excel must be installed. The tool launches a hidden Excel instance, opens the workbook, walks all 31 mapped worksheets, reads every input cell (skipping formulas), validates through Pydantic models, and writes a JSON file. The process takes roughly 21 seconds for a full PHPP workbook.

The resulting JSON is organized by worksheet key. Each worksheet contains some combination of scalar fields, config values, and sections. A section may be a flat dict of key-value pairs, a list of row dicts (for repeating blocks like window schedules), or a grid of column × row intersections.

#### Writing into a blank template

```bash
phpp-tool write records/my_building.json Data/Empty_IP.xlsx -o output.xlsx
```

The tool copies the template, opens the copy in a hidden Excel instance to resolve all cell addresses (label searches, named ranges, block start rows), collects writes, closes Excel without saving, and persists the values via `surgical_writer.py`'s ZIP/XML patch. This takes roughly 48 seconds (dominated by Excel launch and address resolution; persistence itself is fast).

The output file is a valid .xlsx that contains all the designer's input values in the correct cells, with `<extLst>` extensions (Data Validation, etc.) and `<headerFooter>` content preserved byte-for-byte from the template — verified across all 83 sheets of a full roundtrip write. The file still cannot be reliably reopened by Excel via AppleScript automation on macOS 26 (a separate OS/Excel-version compatibility issue, unrelated to the write mechanism), but it is a faithful copy of the template with only the intended cell values changed.

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

Produced by `phpp-tool write`. This is a copy of the blank template with input values injected at the addresses the writer resolved. Formulas remain intact as they were in the template, and `<extLst>`/`<headerFooter>` content is preserved byte-for-byte since persistence goes through `surgical_writer.py`'s ZIP/XML patch rather than an openpyxl save. The file still cannot be reliably reopened by Excel via AppleScript automation on macOS 26 (a separate OS/Excel-version issue), but it can be opened manually in Excel with no degraded features.

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

---

## Part 5 — Concerns, Features, and Limitations

Lessons from building both PHX_xlwg and its sibling PHX_pyxl against the same field map, with xlwings and openpyxl respectively.

### `phpp-field-mapping.md`

**Features**

- Single shared dictionary — both PHX_xlwg and PHX_pyxl read the exact same markdown file and get identical locator behavior, so the field map only needs maintaining once.
- Human-readable and git-diffable — field map edits review like code changes, not like an opaque binary spreadsheet-to-spreadsheet mapping.
- Six addressing strategies (label-anchored, header+entry block, named ranges, absolute address, column+row-offset, fixed result) cover PHPP's inconsistent per-sheet layout without needing sheet-specific code in the reader/writer.
- `phpp-tool inspect-map` gives instant field/section/config coverage auditing after any field map edit.

**Concerns / Limitations**

- ~~**SI/IP unit mismatch**~~ — **Fixed 2026-07-01.** The field map is now versioned: `phpp-field-mapping/EN_10_6_IP.md` and `EN_10_6_SI.md`, selected via `--phpp-version` (default `EN_10_6_IP`). `prefer_si_sheet()`'s runtime "try `<Name> SI` first" guessing is deleted entirely — each version declares its own correct `sheet_name` and locator strings directly. This turned out to matter more than a label mismatch: in an IP-shell workbook the `<Name> SI` tabs are formula mirrors of the base tab's real inputs (see the new item below), so the old default silently lost data, not just risked a wrong unit label. Verified against genuinely distinct files: `Data/Example_IP.xlsx` (IP-shell, dual-tab) and `Data/Example_SI.xlsx` (SI-native, single-shell, no `SI`-suffixed tabs at all) — full details and evidence in `phpp-concerns-and-examples.md` #1.
- **Depends on German internal Excel defined names** (e.g. `Werte_Klima_Region`) — fragile if a PHPP version localizes differently or renames internal ranges.
- **Assumes a stable PHPP layout** — absolute-address and fixed-result strategies hardcode row/column numbers. A PHPP template revision that inserts or removes rows silently breaks these locators with no built-in detection or warning.
- **Stale relative to test workbooks** — the label `"DHW circulation pipes or, for heat interface units, forward and return flows"` is defined in the map but not found in either test PHPP workbook, suggesting the map has drifted from the PHPP versions actually in use.
- ~~**Conflates inputs and outputs**~~ — **Fixed 2026-07-01.** Label-anchored entries can now carry an optional `` `key` (input) ``/`` `key` (output) `` tag, cross-checked against actual cell formula status at read time (warns on drift) and enforced at write time (`writer.py` refuses to write to `(output)`-tagged fields). A one-time migration auto-tagged all 20 label-anchored entries in both `EN_10_6_IP.md`/`EN_10_6_SI.md` against the blank templates. Surprising result while building this: the concern's own original two examples (`phi_building_category_type`, `setpoint_winter`) no longer reproduce — both are literal on the base `Verification` tab, and were only "formula-driven" because of the `<Name> SI` mirror-tab bug below (now fixed). All 20 tags ended up `(input)`, none `(output)` — see `phpp-concerns-and-examples.md` #5.

**Specific data-quality defects (verified 2026-07-01):**

- **Malformed `phi_certification_class` row** (Verification, field-map line 21) — unescaped `|` characters inside the label and options text make this a non-standard markdown table row. It doesn't currently break anything: `map_parser.py`'s `_parse_label_row()` uses a state machine specifically hardened for this row (see its docstring), and the parsed `locator_string` — `"Class | Primary energy method"` — matches the real label text in `Verification SI!T13`. Still, it's fragile: any future change to the label-row parser that doesn't account for embedded pipes would silently break this field.
- **`energy_unit: KHW` typo** (SolarDHW config, field-map line 680) — should be `KWH`, as used everywhere else including the structurally identical PV config block two sections later. Confirmed harmless: `energy_unit`/`footprint_unit` are never read by any Python module — they're descriptive metadata only, not consumed by the reader, writer, or models.
- ~~**Climate `ud_block` header locator is broken**~~ — **Fixed 2026-07-01, and turned out to be a code bug, not bad data.** `PH-Tools/PHX`'s own shape file has the identical "swapped-looking" `header_locator` shape for this section, which was the tell. The real bugs: `start_row: 67` wasn't recognized as an alias for `entry_row_start`/`entry_start_row`, so it was silently ignored; and `_read_section()`'s dispatch order checked `has_items` before checking for `column_fields` anchored by an explicit start row, misrouting to `_read_items_section()`. Fixed both (alias added, dispatch reordered) in `reader.py`/`writer.py` in both PHX_pyxl and PHX_xlwg. `CLIMATE.ud_block` now resolves real monthly climate data — see `phpp-concerns-and-examples.md` #8.
- **Duplicate target cells — one confirmed intentional, one still a genuine bug.** `psi_g_left`/`psi_g_right`/`psi_g_bottom`/`psi_g_top` (Windows → frames, lines 290–293) all mapping to column `IR` is **documented as intentional** in `PHX_Dev/CLAUDE.md` — the original prototype's planning doc — since PHPP treats the glazing-edge (spacer) thermal bridge as one uniform value per window, unlike the installation thermal bridge (`psi_i`), which genuinely varies by side. `duct_assign_1`–`8` (Ventilation, lines 477–484) mapping sequentially to columns Q–X, then `duct_assign_9`/`_10` (lines 485–486) both jumping to `Z` (skipping `Y`) has no such documented rationale and remains a genuine copy-paste-style defect — every duct row silently loses whatever distinct value actually lives in column `Y`.

**Structural concerns affecting efficiency/correctness (verified 2026-07-01):**

- ~~**Config value type inferred from string shape, not declared**~~ — **Fixed 2026-07-01.** Every config/items bullet now requires an explicit `(literal)`/`(address)`/`(named_range)` tag; `map_parser.py` raises `FieldMapError` at parse time if one is missing or invalid, rather than falling back to shape-based guessing. `classify_item()`, `_is_cell_ref()`, and `_is_named_range()` are deleted entirely from `locators.py` — the parsed tag is looked up directly at read/write time. A one-time migration script tagged all 170 existing entries from their then-current classification; the 2 known-wrong `footprint_unit` entries were hand-corrected to `(literal)`. `SOLAR_DHW.footprint_unit`/`SOLAR_PV.footprint_unit` now read back as `"M2"` — see `phpp-concerns-and-examples.md` #10.
- **`UVALUES` and `EASY_PH` are dead worksheet entries.** Their `sheet_name` values (`"U-Values"`, `"easyPH"`) don't match any real sheet in either test workbook — the actual sheets are `"U-values SI"`/`"R-Values"` (different capitalization and naming scheme entirely), and there's no `easyPH` tab at all in PHPP 10.6. Both are silently skipped on every read and write (an INFO-level log line easy to miss in normal output). The commonly quoted "31 mapped worksheets" figure is optimistic — at least 2 of them currently do nothing.
- ~~**The `options` enum metadata is mostly inert.**~~ — **Fixed 2026-07-01.** `reader.py` now checks every resolved label-anchored value's leading code against its field's documented `options` dict and logs a warning on mismatch. Fires correctly on the genuine pre-existing drift in `phi_building_category_type` and stays silent for fields whose resolved code matches — see `phpp-concerns-and-examples.md` #12.
- ~~**Minor: the field map is re-parsed from scratch on every call.**~~ — **Fixed 2026-07-01.** `parse_field_map()` now caches by resolved path + mtime (identical to PHX_pyxl — this module is a direct copy between the two projects) — see `phpp-concerns-and-examples.md` #13.

**Additional structural concerns (verified 2026-07-01, second pass):**

- ~~**`ADDNL_VENT` is silently dropped in its entirety under the default `skip_formulas=True` mode**~~ — **Fixed 2026-07-01. Was the most severe defect found in the first verification pass.** `resolve_block()`'s sparse-row heuristic discarded a row if it had no string value and very few non-`None` fields — meant to detect genuinely blank template rows. Under `skip_formulas=True`, formula-driven fields like `display_name` got nulled out *before* this check ran, tripping the same heuristic on real, populated room rows; two levels up, bare `if result:` truthiness checks then dropped the whole worksheet key silently. **The fix:** sparseness is now decided against raw, unfiltered values first (adapted here to xlwings' batch-read style — the whole region is read in one AppleScript call, so the raw-vs-filtered split happens per-cell within that batch), with `skip_formulas` applied only to what's actually returned; also added a diagnostic warning when a mapped worksheet resolves to no data at all. `read_phpp('Data/Example_IP.xlsx', ...)['ADDNL_VENT']['rooms']` now returns real populated rows, matching PHX_pyxl's result exactly — see `phpp-concerns-and-examples.md` #14.
- ~~**`entry_row_start`, when present, silently overrides the discovered entry-locator row with no cross-check.**~~ — **Fixed 2026-07-01.** Both `resolve_block()` and `_read_column_row_section()` (the `tanks` section actually dispatches through the latter, not the former) still use `entry_row_start` when present, but now also discover the label's actual row and log a warning when the two disagree. The `tanks` section's own drift (`entry_row_start: 191` vs. label `"Storage type 1"` actually at row 189) is real and still present in the field map — that's a data question left to field-map maintenance, not something the code fix should silently resolve one way or the other. **Building this fix directly led to finding #26 below** — the first real exercise of `find_row_in_col()` on this code path surfaced a pre-existing, more serious bug in how this project batch-reads Excel ranges. See `phpp-concerns-and-examples.md` #15.

**New findings and versioning architecture (2026-07-01, third pass — see `phpp-concerns-and-examples.md` #24-25 for full evidence):**

- ~~**`<Name> SI` tabs in IP-shell workbooks are formula mirrors of the base tab, not independent data**~~ — **Fixed 2026-07-01.** Sampling formula-vs-value ratios across all 28 worksheet pairs in `Example_IP.xlsx` showed a minority of worksheets where the SI tab has noticeably more formula cells than the base tab — exactly where genuine input cells live on the base tab and get mirrored onto the SI tab via `=IF(...)`-style passthrough formulas for unit-converted display. Since `skip_formulas` treats any formula cell as "not a real input," the old `prefer_si_sheet()` default silently discarded real designer input workbook-wide, not just in `ADDNL_VENT`. Fixed by deleting `prefer_si_sheet()` and moving to per-version explicit sheet names (see the SI/IP concern above).
- ~~**`Climate.ud_block`'s `summer_delta_t_unit` had a bogus `DELTA-C` "column" value**~~ — **Fixed 2026-07-01.** Every sibling row in that column-fields table maps to a real 1-3 letter column; this one had the literal string `DELTA-C` — a stray unit annotation in the wrong table cell. Harmless on PHX_pyxl (`col_to_idx("DELTA-C")` silently produces a nonsensical but non-crashing index) but **crashed this project's live AppleScript call outright** — `appscript.reference.CommandError: ... columns[1300905401].get_address()` — when discovered during the Stage D port from PHX_pyxl. Removed from all four field map copies (`EN_10_6_SI.md`/`EN_10_6_IP.md` × PHX_pyxl/PHX_xlwg) — it never represented a real column mapping and nothing consumed the key. This is a good example of why testing a shared field map against both backends matters: the same bad data was silent on one and fatal on the other.
- **Field map is now versioned.** `phpp-field-mapping/EN_10_6_IP.md` and `EN_10_6_SI.md` replace the single `phpp-field-mapping.md`, selected via `--phpp-version` (default `EN_10_6_IP`). Adopted from the architectural pattern in `PH-Tools/PHX`'s `phpp_localization/` directory (one shape file per language/version/unit-variant) — though nothing was copied from that GPL-3.0 project; both files here are independently authored and verified against this project's own `Example_IP.xlsx`/`Example_SI.xlsx`.

### xlwings

**Features**

- Drives a live Excel instance — formulas always recalculate, so reads never suffer from cached-value staleness the way a pure-openpyxl read can.
- No file-format degradation during read or address resolution — Excel serializes its own file, so charts, drawings, and data validations survive untouched while xlwings is just locating cells.
- Batch reads (`find_row_in_col()`, `resolve_block()`) fetch a column or region in row-chunks rather than cell-by-cell, mitigating per-call RPC latency while staying correct near hidden-row boundaries (see the fixed concern below — this used to be a single unchunked call per read).
- The write path resolves addresses with the exact same locator code and live data as the read path, so label searches, named ranges, and entry-row detection behave identically in both directions.

**Concerns / Limitations**

- **Hard Excel dependency for every operation** — read, the write path's address-resolution phase, and the entire test suite (~9s) all require a running Excel instance; there is no headless or CI-friendly mode. This is exactly the constraint PHX_pyxl was built to eliminate.
- **macOS 26 broke every native AppleScript save path** for large workbooks: `wb.save()` returns error -50, `close(saving=yes)` hangs indefinitely, and even a VBA `ActiveWorkbook.Save` macro errors -50. Small test fixtures save fine; production-size PHPP workbooks do not. This is why the writer can't be pure xlwings — it must resolve addresses live, then quit Excel and persist via `surgical_writer.py` instead.
- ~~**The file-integrity advantage only covers the read side**~~ — **No longer true as of 2026-07-01.** Persistence used to go through openpyxl's `save()`, which inherited openpyxl's extension-dropping behavior regardless of how cleanly xlwings resolved the addresses. Now that persistence goes through `surgical_writer.py`'s ZIP/XML patch instead, the file-integrity advantage covers both the read/resolution side (Excel never touches the file format) and the write side (the patch never invokes openpyxl's serializer).
- **Slower per operation** — Excel process launch and AppleScript RPC overhead make reads (~21s) and writes (~48s) noticeably slower than the openpyxl-only equivalents (~20s / ~30s) in PHX_pyxl.
- **Operational fragility** — transient AppleScript error -50 on rapid open/close cycles (mitigated with a one-retry-after-delay in `open_book()`), alert dialogs that must be suppressed, and Excel processes that can be left running if a script exits abnormally.
- ~~**Batch reads could silently drop a row's worth of data near hidden-row boundaries.**~~ — **Fixed 2026-07-01. Found while implementing the `entry_row_start` cross-check above** — that fix called `find_row_in_col()` for real on a code path that never exercised it before, and the discovered row (188) didn't match PHX_pyxl's answer (189) for the identical file. Bisected the cause precisely: requesting a column's values for rows 1–132 returned exactly 132 values, but requesting rows 1–133 (one more row, hidden, `row_height == 0.0`) also returned only 132 values, with no error — the hidden row's value vanished from the array, shifting every subsequent row index by one. This affected both `find_row_in_col()`'s label search and `resolve_block()`'s actual field-data reads (the same underlying single-shot `.value`/`.formula` batch call), so the exposure wasn't limited to one field — it's a property of xlwings' Mac AppleScript backend, not of any particular locator. **The fix:** both functions now read in 50-row chunks via a shared `_read_rect_chunked()` helper that validates each chunk's returned row count against what was requested; a chunk that comes back short falls back to a per-row read for just that chunk. Verified: `find_row_in_col()` now returns 189 (matching PHX_pyxl exactly) on both `Empty_IP.xlsx` and `Empty_SI.xlsx`, and full roundtrips of both variants still show a perfect cell-by-cell match with no measurable timing regression. **Caution learned while diagnosing this:** don't reach for `.api` (COM-style property access) to inspect row visibility on this Mac AppleScript backend — it isn't supported the same way as on Windows, and calling it crashed the underlying Excel process outright (triggering macOS's crash reporter) rather than raising a catchable Python exception. `.row_height` (used in the final diagnosis and safe here) is the right tool for this. See `phpp-concerns-and-examples.md` #26.

### openpyxl

**Features**

- No Excel needed for locator resolution or the roundtrip test's final verification step — reading the written file back with openpyxl to confirm every writer-reported write actually landed at the correct cell.
- Keeps the address-resolution half of the writer small — because persistence now lives entirely in `surgical_writer.py`, `writer.py` itself only ever collects a pending list of `(sheet, col, row, value)` tuples; ~290 lines, versus ~770 lines in the original all-openpyxl PHX_Dev prototype that traced and rewrote formula chains directly.

**Concerns / Limitations**

- ~~**Used to drop content on save, via two distinct mechanisms, even though addresses were resolved live in Excel moments earlier**~~ — **Fixed 2026-07-01.** This applied when persistence went through openpyxl's `save()`: (1) `parse_extensions()` unconditionally discards any of 8 GUID-tagged `<extLst>` extension types it recognizes — Conditional Formatting, Data Validation, Sparkline Group, Slicer List, Protected Range, Ignored Error, Web Extension, Timeline Ref (`openpyxl/xml/constants.py:EXT_TYPES`) — of which only **Data Validation** was confirmed to fire on `Example.xlsx`/`Empty.xlsx`. (2) `header_footer.py`'s `_split_string()` silently blanks all three header/footer sections when PHPP's actual string doesn't match its `&L...&C...&R...` pattern. The hybrid writer now persists via `surgical_writer.py`'s ZIP/XML patch instead of openpyxl's `save()`, touching only the `<sheetData>` region of affected sheets. Verified byte-for-byte: `<extLst>` and `<headerFooter>` regions are identical between template and written output across all 83 sheets of a full 13,102-cell roundtrip write — the file-integrity advantage now genuinely covers both the resolution *and* the persistence side.
- **No recalculation** — the written file's formula results won't refresh until it is manually opened and saved in Excel, the same manual-step requirement PHX_pyxl documents for its `verify_excel.py` full-fidelity check.
- **Written files generally can't be reliably reopened by Excel via AppleScript** afterward (validation-error hangs) — this is unrelated to file integrity (the content itself is intact and complete, verified byte-for-byte), but it does mean the pipeline can't fully automate a round trip through live Excel; opening the output requires a manual step.
