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
├── pyproject.toml               ← deps: xlwings, click, pydantic, openpyxl, lxml
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
# Install in dev mode
pip install -e ".[dev]"

# Run tests
pytest tests/ -v

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
