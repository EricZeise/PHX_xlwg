# CLAUDE.md — PHPP Data Tool (xlwings backend)

## What this project is

A Python CLI that **reads designer-entered input data from a filled PHPP workbook** (Passive House Planning Package), **stores it as a portable JSON record**, and **writes that record back into a blank PHPP workbook**.

This is a rewrite of the openpyxl-based PHX_Dev tool using **xlwings** as the Excel backend. xlwings drives a live Excel instance, which means:
- **Formulas recalculate automatically** — no cached-value issues
- **No formula chain tracing** — write to input cells, formulas update
- **Charts, drawings, data validations all survive** — writes persist via a surgical ZIP/XML patch (same mechanism PHX_Dev originated), not an openpyxl save, so `<extLst>`/`<headerFooter>` content is preserved

### MVP pipeline

```
Filled PHPP (.xlsx)  →  read  →  JSON record  →  write  →  Blank PHPP (.xlsx)
```

### Requirement: local Excel installation

xlwings requires Excel (macOS or Windows). This is acceptable because PHPP designers already have Excel installed. A headless Linux deployment would need the openpyxl fallback (see PHX_Dev).

---

## Architecture

```
phpp-field-mapping/EN_10_6_IP.md   (locator dictionary — where each field lives, versioned by PHPP type)
        ↓
   map_parser.py        (parse markdown → structured dict; requires explicit type tags)
        ↓
   locators.py          (6 addressing strategies to find cells)
        ↓
 ┌──────┴──────┐
 reader.py    writer.py  (xlwings locators + surgical XML persistence)
 └──────┬──────┘
     models.py           (pydantic validation)
        ↓
     cli.py              (Click CLI, --phpp-version selects the field map)
```

### Versioned field maps

`phpp-field-mapping/` holds one field map per PHPP workbook type — `EN_10_6_IP.md` (IP-shell workbook, which also carries `<Name> SI`-suffixed mirror tabs) and `EN_10_6_SI.md` (genuinely SI-native, single-shell workbook). Selected via `--phpp-version` (default `EN_10_6_IP`). There is no runtime SI/IP sheet-name guessing — each version declares its own correct `sheet_name` directly. This matters because in an IP-shell workbook, the `<Name> SI` tabs are formula mirrors of the base tab's real input cells, not independent data — reading them under `skip_formulas` used to silently lose designer input.

### What changed from PHX_Dev (openpyxl)

| Concern | openpyxl (PHX_Dev) | xlwings (this project) |
|---------|-------------------|----------------------|
| Formula values | Stale cached values; needed `data_only=True/False` | Always current — Excel recalculates |
| Writing to formula cells | Phase 1b: trace formula chains, redirect writes, overwrite formulas | Not needed — write to input cells, formulas update |
| File integrity | Surgical ZIP/XML splicing (lxml) | Same surgical ZIP/XML splicing (`surgical_writer.py`), fed by xlwings-resolved addresses instead of openpyxl-resolved ones |
| Dependencies | openpyxl, lxml | xlwings, openpyxl, lxml (requires Excel for reads + locator resolution) |
| writer.py size | ~770 lines | ~290 lines (address resolution only; persistence lives in `surgical_writer.py`) |
| Headless operation | Yes | No (needs Excel) |

### What stayed the same

- `map_parser.py` — identical (pure text parsing, no Excel dependency)
- `models.py` — identical (pure pydantic)
- `locators.py` — same 6 strategies, xlwings cell access instead of openpyxl
- `cli.py` — same interface
- `phpp-field-mapping/` — same field maps, kept as separate synced copies in each repo
- `compare_json/` — same comparison tool
- All test fixtures for map_parser, models, CLI

---

## Repository structure

```
PHX_xlwg/
├── CLAUDE.md                    ← this file
├── pyproject.toml               ← deps: xlwings, click, pydantic, openpyxl, lxml; Ruff / isort / ty / pytest settings
├── uv.lock                      ← locked development environment (`uv sync`)
├── phpp-field-mapping/           ← versioned locator dictionaries (31 worksheets each)
│   ├── EN_10_6_IP.md             ← IP-shell workbook (base tabs + SI mirror tabs)
│   └── EN_10_6_SI.md             ← genuinely SI-native, single-shell workbook
├── src/
│   ├── phpp_tool/
│   │   ├── __init__.py
│   │   ├── cli.py               ← Click CLI: read / write / inspect-map, --phpp-version
│   │   ├── map_parser.py        ← Parse a field map file → structured dict; requires type tags
│   │   ├── locators.py          ← 6 addressing strategies (xlwings)
│   │   ├── reader.py            ← xlwings-based reader
│   │   ├── writer.py            ← xlwings address resolution, collects writes (~290 lines)
│   │   ├── atomic.py            ← Atomic file output (temp file beside the destination, then replace)
│   │   ├── surgical_writer.py   ← Persists writes via ZIP/XML patch (lxml), preserving extLst/headerFooter
│   │   ├── excel_app.py         ← Excel app factory (macOS path detection)
│   │   └── models.py            ← Pydantic models for building record JSON
│   └── compare_json/            ← Standalone JSON diff tool
├── tests/
└── records/                     ← Output directory for JSON building records
```

---

## Commands

```bash
# Set up the development environment (uv: pyproject.toml + uv.lock; installs the package editable plus
# pytest, Ruff, isort, ty and lxml-stubs). `pip install -e ".[dev]"` still works without uv.
uv sync

# Run the tests that need no Excel (104 tests, <1 second; this is what the review gate runs)
uv run pytest -m "not excel" -v

# Run everything, including the live-Excel roundtrip (launches Excel 2021 hidden; macOS shows
# file-access consent prompts -- see tests/test_cli.py)
uv run pytest tests/ -v

# The review gate (PH-Demonstrator's CodeReview.md, Step 1) -- all must pass before a commit
uv run ruff check . && uv run ruff format --check . && uv run isort --check-only . \
  && uv run ty check --error-on-warning && uv lock --check && uv run pytest -q -m "not excel"

# Step 2 (detailed review): report-only counts may only go down; coverage may not drop
uv run python scripts/check_ratchet.py && uv run python scripts/run_with_coverage.py

# Read a PHPP into JSON (--phpp-version defaults to EN_10_6_IP)
phpp-tool read Data/Example_IP.xlsx -o records/my_building.json --phpp-version EN_10_6_IP

# Read a genuinely SI-native PHPP
phpp-tool read Data/Example_SI.xlsx -o records/my_building.json --phpp-version EN_10_6_SI

# Write a record into a blank PHPP -- must match the --phpp-version used for read
phpp-tool write records/my_building.json Data/Empty_IP.xlsx -o output.xlsx --phpp-version EN_10_6_IP

# Inspect the field map
phpp-tool inspect-map --phpp-version EN_10_6_IP
```

---

## Constraints

- **Never read or embed PHPP formulas** — only designer-entered input values.
- **The field map is the single source of truth** for cell locations.
- **Requires Excel** — xlwings drives a live Excel instance.
- **Graceful degradation**: missing sheets, empty cells, and unresolvable locators produce warnings, not crashes.
- **Read and write must use the same `--phpp-version`** for a given workbook — `read` stamps the output JSON with `_phpp_version`; `write` warns (doesn't hard-block) if it doesn't match.
- **Every config/items field-map entry requires an explicit type tag** (`(literal)`/`(address)`/`(named_range)`) — `map_parser.py` raises `FieldMapError` at parse time if one is missing or invalid.

---

## Review gate (2026-11-09)

PH-Demonstrator's Step 1 code-review gate is applied here, with the Excel dependency in mind: `pyproject.toml` carries the tool settings (line length 88, target Python 3.10 = the minimum supported; Ruff rules E, W, F, B, UP, C4, PIE, PLW, PLE; isort profile black; ty at Python 3.10 with a `tests/**` override that relaxes `unresolved-attribute`, `invalid-argument-type` and `not-subscriptable` because the tests read Optional sections of the pydantic record directly) and `uv.lock` the development environment (`.python-version` 3.13; `lxml-stubs` because `lxml.etree` is compiled). **Nothing in the gate launches Excel:** the one live-Excel test (`TestExcelRoundTrip`) is marked `excel`, the gate runs `pytest -m "not excel"`, and the shared-Excel fixture in `tests/conftest.py` is no longer autouse -- it starts Excel only for a test that asks for it (before, every pytest run launched Excel and could trigger macOS consent prompts even for pure unit tests). **xlwings is pinned to 0.36.8** (`constraint-dependencies`): the AppleScript backend is version-sensitive and the gate cannot exercise it, so a bump is deliberate and must be followed by the `excel` tests and `scripts/roundtrip.py`. The old `.venv` had its editable install pointed at a path from before the repo moved under `PHX/`; `uv sync` rebuilt it. Fixes made: unused imports and f-strings removed, one overwritten loop variable renamed (`map_parser._parse_options`), `excel_app.open_book(**kwargs: Any)`, 16 files reformatted -- behaviour unchanged. `excel_app._find_excel()` still passes an explicit Office 2021/2019/2016 path (never a bare `xw.App()`), and nothing here calls `.api`.

**Atomic output (same day, ported from PHX_pyxl):** the JSON building record (`phpp-tool read -o`), the `compare_json` reports and the surgical workbook rewrite (`surgical_writer._rebuild_zip` / the no-writes copy) go through `phpp_tool/atomic.py` (temp file beside the destination, `fsync`, `os.replace`; `tests/test_atomic.py`). **Known limit, not changed:** `writer.write_phpp` still starts by copying the template onto the output path (`shutil.copy2`, then `chmod`) *before* Excel is opened, so an existing output file is replaced at that point and a failure between that copy and the final atomic rewrite (Excel hang, user cancel) leaves a template copy rather than the previous output. Closing it means opening a working copy in Excel instead of the output path; that changes the live-Excel flow (file-access prompts, lock files), so it needs a run with Excel before it is trusted.

---

## Step 2 (2026-11-09, later)

The detailed review of PH-Demonstrator's `CodeReview.md` was run: Ruff also enforces SIM, PERF, FURB, RUF, ARG, S, T20, N and BLE (ignored on purpose: `S101` -- the tests' asserts; `cli.py` and the scripts may print; scripts / tests may run subprocesses). Findings fixed: nested `if` / `with` combined, `x if x else None` -> `x or None`, three `append` loops -> `extend`, `[*a, b]` for `a + [b]`, two unused parameters handled (`reader._read_section(ws_config)` removed, `_read_appliance_section(ws)` renamed `_ws` -- it is a stub); `airtightness_Vn50` kept (the PHPP field's own name, `noqa: N815`); two `noqa: BLE001` with reasons (`excel_app.open_book`'s retry, the Excel fixture's skip). **These edits touch `reader.py` and `writer.py`, which only run against live Excel; they are mechanical and the unit tests pass, but the `excel` test and `scripts/roundtrip.py` have not been re-run -- run them once before trusting a release.** `scripts/check_ratchet.py` holds the report-only `C90` / `PLR` counts and the `noqa` / `ty: ignore` counts against `scripts/ratchet_baseline.json` (90 findings, 3 `noqa`); `scripts/run_with_coverage.py` runs pytest (`-m "not excel"`) under coverage (7.17.2, Python 3.11+ only) against `scripts/coverage_baseline.json` (43.34 % statements + branches -- the Excel-driven reader, writer and locators are only covered by the live test). `uv lock --upgrade --dry-run` showed nothing newer except the deliberately pinned xlwings, and `pip-audit` found no known vulnerabilities. Not done: the complexity refactors, an Excel-free fake of the xlwings objects to unit-test the reader / writer, CI, Python 3.10 verification.

