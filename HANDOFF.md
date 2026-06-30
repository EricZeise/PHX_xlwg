# PHX_xlw Handoff — 2026-06-29

## Current state

**All 91 tests pass.** Reader works end-to-end. Writer works for small workbooks (test fixtures) and completes on the real 83-sheet PHPP in ~41 seconds. No git repo initialized yet — all files are untracked.

## What was built

A rewrite of PHX_Dev's openpyxl-only pipeline using xlwings for reading and locator resolution. See `CLAUDE.md` for full architecture.

### Key files

| File | Role |
|------|------|
| `src/phpp_tool/reader.py` | xlwings-based reader — opens Excel, reads cells, returns dict |
| `src/phpp_tool/writer.py` | Hybrid writer — xlwings resolves cell addresses, openpyxl persists values |
| `src/phpp_tool/excel_app.py` | Excel app factory — finds correct Excel installation on macOS, `_close_saving()` fallback chain |
| `src/phpp_tool/locators.py` | 6 addressing strategies (label-anchored, block, row-offset, column-row, static-column, items) |
| `tests/test_cli.py` | 8 tests including full read→JSON→write→verify roundtrip |
| `tests/test_models.py` | 17 tests including xlwings integration test |
| `tests/test_map_parser.py` | 66 tests for field map parsing |

### The macOS 26 save problem and the hybrid writer

On macOS 26 + Excel 2021, **all AppleScript save paths are broken for large workbooks:**
- `wb.save()` → error -50 (parameter error)
- `close(saving=yes)` → hangs indefinitely (even with `calculation="manual"`, even with `timeout=-1`)
- `run VB macro string "ActiveWorkbook.Save"` → error -50

Small workbooks (1–2 sheets) save fine via `close(saving=yes)`.

**Solution in `writer.py`:** xlwings opens the PHPP with `calculation="manual"` to resolve locator lookups (label searches, named ranges, entry-row detection). `_write_cell()` does NOT write to xlwings — it appends `(sheet_name, col, row, value)` tuples to `_pending_writes`. After `wb.close()` + `app.quit()`, `_apply_writes_openpyxl()` opens the file with openpyxl and persists all collected values.

`_close_saving()` in `excel_app.py` is still available for small-workbook callers (tests use it implicitly via the test fixtures). It tries `wb.save()` → `close(saving=yes, timeout=60)` → `wb.close()` (no save).

### Environment details

- macOS 26.5.1, boot volume named "System"
- Two Excel versions installed: Office 2011 (`/Applications/Microsoft Excel.app`) and Office 2021 (`/Applications/Microsoft Office 2021/Microsoft Excel.app`)
- `excel_app()` uses xlwings `spec` parameter to select Office 2021
- xlwings uses appscript (aeosa) under the hood — HFS paths become `System:Users:...`
- Python 3.13.3, venv at `.venv/`
- `Data/` is a symlink to `../PHX_Dev/Data/` — contains Empty.xlsx, Example.xlsx, Model.xlsx

## Roundtrip test results (from this session)

Tested against real PHPP workbooks in `Data/`:
- **Empty→Empty roundtrip** (xlwings read → JSON → openpyxl write → xlwings read): 1 difference (trailing space) — 99.99% match
- **Example→Empty roundtrip**: 60 differences out of ~20,000 cells — 99.7% match (same as PHX_Dev baseline)
- Roundtrip artifacts are in `records/roundtrip_20260629_190740/`

## Known issues / next steps

1. **Field map labels assume SI units (`[°C]`)** — the PHPP Empty.xlsx uses IP units (`[°F]`), so label-anchored locators like `setpoint_winter` don't match. This affects writes to sheets where the locator string contains unit text. The reader has the same issue but reads successfully because those cells have values regardless. Fix: either add IP-unit variants to the field map, or strip unit suffixes during label matching.

2. **`_pending_writes` is a module-level mutable list** — fine for single-threaded CLI use but not thread-safe. If this ever needs to be called concurrently, refactor to pass the list through the call chain.

3. **openpyxl write may strip some Excel features** — openpyxl's `load_workbook` + `save` cycle can lose Excel-specific features (sparklines, custom XML, some data validations). For the PHPP writer this is acceptable because we only write to input cells, but files saved this way may trigger Excel's "recovery" dialog on first open. The PHX_Dev surgical writer (lxml-based ZIP splicing) avoids this but is far more complex.

4. **No git repo yet** — consider `git init` and initial commit.

5. **`records/roundtrip_20260629_190740/` contains test artifacts** — can be cleaned up or gitignored.

6. **`_close_saving()` is no longer used by `writer.py`** — it's only used by test fixtures (small workbooks). Could be removed or kept for Windows compatibility where `wb.save()` works fine.
