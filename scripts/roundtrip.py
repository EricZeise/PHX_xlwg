#!/usr/bin/env -S python3 -u
"""Roundtrip test: read inputs → JSON → write → verify every written cell.

Usage:
    python scripts/roundtrip.py Data/Example.xlsx Data/Empty.xlsx
    python scripts/roundtrip.py Data/Empty.xlsx Data/Empty.xlsx

First argument is the source workbook to read.
Second argument is the blank template to write into.
Results are saved to records/roundtrip_<timestamp>/.

Phase 1 (read inputs):  xlwings reads only input cells (skip_formulas=True).
Phase 2 (read all):     xlwings reads ALL cells (skip_formulas=False) —
                        shows what the formula filter removed.
Phase 3 (write):        xlwings locators + openpyxl persistence.  The writer
                        returns its list of (sheet, col, row, value) writes.
Phase 4 (verify):       openpyxl reads the written file at every address the
                        writer targeted and compares cell-by-cell.
"""

from __future__ import annotations

import json
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from phpp_tool.locators import col_to_idx
from phpp_tool.reader import read_phpp
from phpp_tool.writer import write_phpp

FIELD_MAP = str(ROOT / "phpp-field-mapping.md")


def _count_values(data: dict, depth: int = 0) -> tuple[int, int]:
    """Count (total_values, non_none_values) recursively."""
    total = 0
    non_none = 0
    for v in data.values():
        if isinstance(v, dict):
            t, n = _count_values(v, depth + 1)
            total += t
            non_none += n
        elif isinstance(v, list):
            for item in v:
                if isinstance(item, dict):
                    t, n = _count_values(item, depth + 1)
                    total += t
                    non_none += n
                else:
                    total += 1
                    if item is not None:
                        non_none += 1
        else:
            total += 1
            if v is not None:
                non_none += 1
    return total, non_none


def _values_match(expected: Any, actual: Any) -> bool:
    """Compare values with tolerance for float rounding."""
    if expected == actual:
        return True
    if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        if abs(expected) < 1e-10 and abs(actual) < 1e-10:
            return True
        if abs(expected) > 0:
            return abs(expected - actual) / abs(expected) < 1e-6
    if isinstance(expected, float) and isinstance(actual, int):
        return _values_match(expected, float(actual))
    if isinstance(expected, int) and isinstance(actual, float):
        return _values_match(float(expected), actual)
    return False


def _verify_writes(
    written_path: Path,
    writes: list[tuple[str, str, int, Any]],
) -> tuple[int, int, list[str]]:
    """Verify every cell the writer targeted.

    Opens the written file with openpyxl and checks each
    (sheet, col, row, value) against what was actually persisted.
    Returns (checked, mismatches, detail_lines).
    """
    import warnings
    warnings.filterwarnings("ignore", category=UserWarning)
    from openpyxl import load_workbook

    wb = load_workbook(str(written_path), data_only=False)
    checked = 0
    mismatched = 0
    details: list[str] = []

    for sheet_name, col, row, expected in writes:
        if sheet_name not in wb.sheetnames:
            mismatched += 1
            details.append(f"  MISSING SHEET: {sheet_name}")
            continue

        ws = wb[sheet_name]
        actual = ws.cell(row=row, column=col_to_idx(col)).value
        checked += 1

        if not _values_match(expected, actual):
            mismatched += 1
            details.append(
                f"  {sheet_name}!{col}{row}:"
                f" expected {expected!r} ({type(expected).__name__}),"
                f" got {actual!r} ({type(actual).__name__})")

    wb.close()
    return checked, mismatched, details


def roundtrip(source: Path, template: Path, out_dir: Path) -> dict:
    source_name = source.stem
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n{'=' * 60}")
    print(f"  ROUNDTRIP: {source.name} → {template.name}")
    print(f"{'=' * 60}")

    # Step 1: read source — input cells only
    print(f"\n[1/4] Reading {source.name} — inputs only (xlwings) ...")
    t0 = time.time()
    data_inputs = read_phpp(str(source), FIELD_MAP, skip_formulas=True)
    t_read_inputs = time.time() - t0
    total_i, non_none_i = _count_values(data_inputs)
    print(f"      {len(data_inputs)} worksheets,"
          f" {non_none_i} non-None / {total_i} total values,"
          f" {t_read_inputs:.1f}s")

    json_inputs = out_dir / f"{source_name}_inputs.json"
    json_inputs.write_text(
        json.dumps(data_inputs, indent=2, default=str), encoding="utf-8"
    )

    # Step 2: read source — all cells (for formula filter stats)
    print(f"\n[2/4] Reading {source.name} — all cells (xlwings) ...")
    t0 = time.time()
    data_all = read_phpp(str(source), FIELD_MAP, skip_formulas=False)
    t_read_all = time.time() - t0
    total_a, non_none_a = _count_values(data_all)
    print(f"      {len(data_all)} worksheets,"
          f" {non_none_a} non-None / {total_a} total values,"
          f" {t_read_all:.1f}s")
    filtered = non_none_a - non_none_i
    pct = 100 * filtered / max(non_none_a, 1)
    print(f"      Formula filter removed {filtered} values ({pct:.1f}%)")

    json_all = out_dir / f"{source_name}_all.json"
    json_all.write_text(
        json.dumps(data_all, indent=2, default=str), encoding="utf-8"
    )

    # Step 3: write into template (writer returns its write list)
    written_path = out_dir / f"{source_name}_written.xlsx"
    print(f"\n[3/4] Writing into {template.name} → {written_path.name} ...")
    t0 = time.time()
    writes = write_phpp(data_inputs, str(template), str(written_path), FIELD_MAP)
    t_write = time.time() - t0
    print(f"      {len(writes)} cell writes, {t_write:.1f}s")

    # Step 4: verify every written cell
    print(f"\n[4/4] Verifying {len(writes)} written cells (openpyxl) ...")
    t0 = time.time()
    checked, mismatched, mismatch_details = _verify_writes(
        written_path, writes)
    t_verify = time.time() - t0

    if mismatched == 0:
        print(f"      *** All {checked} cells verified — PERFECT MATCH ***")
    else:
        print(f"      {checked} checked, {mismatched} MISMATCHES:")
        for d in mismatch_details[:20]:
            print(d)
        if len(mismatch_details) > 20:
            print(f"      ... and {len(mismatch_details) - 20} more")
    print(f"      {t_verify:.1f}s")

    total_t = t_read_inputs + t_read_all + t_write + t_verify
    print(f"\n  Timing: inputs {t_read_inputs:.1f}s + all {t_read_all:.1f}s"
          f" + write {t_write:.1f}s + verify {t_verify:.1f}s"
          f" = {total_t:.1f}s total")
    print(f"  Artifacts saved to {out_dir}/")

    return {
        "source": source.name,
        "template": template.name,
        "inputs_non_none": non_none_i,
        "all_non_none": non_none_a,
        "formulas_filtered": filtered,
        "cells_written": len(writes),
        "cells_verified": checked,
        "mismatches": mismatched,
        "time_total": total_t,
    }


def main() -> None:
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    pairs = []
    for i in range(1, len(sys.argv), 2):
        source = Path(sys.argv[i])
        template = Path(sys.argv[i + 1]) if i + 1 < len(sys.argv) else None
        if template is None:
            print(f"Missing template for {source}")
            sys.exit(1)
        pairs.append((source, template))

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = ROOT / "records" / f"roundtrip_{stamp}"

    summaries = []
    for source, template in pairs:
        summaries.append(roundtrip(source, template, out_dir))

    if len(summaries) > 1:
        print(f"\n\n{'=' * 60}")
        print("  COMBINED SUMMARY")
        print(f"{'=' * 60}\n")
        for s in summaries:
            print(f"  {s['source']:20s} → {s['template']:20s}"
                  f"  written: {s['cells_written']:>4}"
                  f"  verified: {s['cells_verified']:>4}"
                  f"  mismatches: {s['mismatches']:>3}"
                  f"  formulas filtered: {s['formulas_filtered']:>4}"
                  f"  time: {s['time_total']:.1f}s")


if __name__ == "__main__":
    main()
