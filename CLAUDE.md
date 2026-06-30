# CLAUDE.md — PHPP Data Tool (xlwings backend)

## What this project is

A Python CLI that **reads designer-entered input data from a filled PHPP workbook** (Passive House Planning Package), **stores it as a portable JSON record**, and **writes that record back into a blank PHPP workbook**.

This is a rewrite of the openpyxl-based PHX_Dev tool using **xlwings** as the Excel backend. xlwings drives a live Excel instance, which means:
- **Formulas recalculate automatically** — no cached-value issues
- **No surgical XML editing** — Excel handles its own serialization
- **No formula chain tracing** — write to input cells, formulas update
- **Charts, drawings, data validations all survive** — Excel manages them natively

### MVP pipeline

```
Filled PHPP (.xlsx)  →  read  →  JSON record  →  write  →  Blank PHPP (.xlsx)
```

### Requirement: local Excel installation

xlwings requires Excel (macOS or Windows). This is acceptable because PHPP designers already have Excel installed. A headless Linux deployment would need the openpyxl fallback (see PHX_Dev).

---

## Architecture

```
phpp-field-mapping.md   (locator dictionary — where each field lives)
        ↓
   map_parser.py        (parse markdown → structured dict)
        ↓
   locators.py          (6 addressing strategies to find cells)
        ↓
 ┌──────┴──────┐
 reader.py    writer.py  (xlwings locators + openpyxl persistence)
 └──────┬──────┘
     models.py           (pydantic validation)
        ↓
     cli.py              (Click CLI)
```

### What changed from PHX_Dev (openpyxl)

| Concern | openpyxl (PHX_Dev) | xlwings (this project) |
|---------|-------------------|----------------------|
| Formula values | Stale cached values; needed `data_only=True/False` | Always current — Excel recalculates |
| Writing to formula cells | Phase 1b: trace formula chains, redirect writes, overwrite formulas | Not needed — write to input cells, formulas update |
| File integrity | Surgical ZIP/XML splicing (lxml) | xlwings for addressing, openpyxl for persistence |
| Dependencies | openpyxl, lxml | xlwings, openpyxl (requires Excel for reads + locator resolution) |
| writer.py size | ~770 lines | ~290 lines |
| Headless operation | Yes | No (needs Excel) |

### What stayed the same

- `map_parser.py` — identical (pure text parsing, no Excel dependency)
- `models.py` — identical (pure pydantic)
- `locators.py` — same 6 strategies, xlwings cell access instead of openpyxl
- `cli.py` — same interface
- `phpp-field-mapping.md` — same field map
- `compare_json/` — same comparison tool
- All test fixtures for map_parser, models, CLI

---

## Repository structure

```
PHX_xlw/
├── CLAUDE.md                    ← this file
├── pyproject.toml               ← deps: xlwings, click, pydantic, openpyxl
├── phpp-field-mapping.md        ← the locator dictionary (31 worksheets)
├── src/
│   ├── phpp_tool/
│   │   ├── __init__.py
│   │   ├── cli.py               ← Click CLI: read / write / inspect-map
│   │   ├── map_parser.py        ← Parse phpp-field-mapping.md → structured dict
│   │   ├── locators.py          ← 6 addressing strategies (xlwings)
│   │   ├── reader.py            ← xlwings-based reader
│   │   ├── writer.py            ← xlwings locators + openpyxl writer (~340 lines)
│   │   ├── excel_app.py         ← Excel app factory (macOS path detection)
│   │   └── models.py            ← Pydantic models for building record JSON
│   └── compare_json/            ← Standalone JSON diff tool
├── tests/
└── records/                     ← Output directory for JSON building records
```

---

## Commands

```bash
# Install in dev mode
pip install -e ".[dev]"

# Run tests
pytest tests/ -v

# Read a PHPP into JSON
phpp-tool read path/to/filled_PHPP.xlsx -o records/my_building.json

# Write a record into a blank PHPP
phpp-tool write records/my_building.json path/to/blank_PHPP.xlsx -o output.xlsx

# Inspect the field map
phpp-tool inspect-map
```

---

## Constraints

- **Never read or embed PHPP formulas** — only designer-entered input values.
- **The field map is the single source of truth** for cell locations.
- **Requires Excel** — xlwings drives a live Excel instance.
- **Graceful degradation**: missing sheets, empty cells, and unresolvable locators produce warnings, not crashes.
