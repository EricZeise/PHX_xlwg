#!/usr/bin/env -S python3 -u
"""Roundtrip test: read → JSON → write → read-back → compare.

Usage:
    python scripts/roundtrip.py Data/Example.xlsx Data/Empty.xlsx
    python scripts/roundtrip.py Data/Empty.xlsx Data/Empty.xlsx

First argument is the source workbook to read.
Second argument is the blank template to write into.
Results are saved to records/roundtrip_<timestamp>/.

Phase 1 (read) uses xlwings+Excel to get live formula values.
Phase 2 (write) uses xlwings for locator resolution + openpyxl for
persistence — the only save path that doesn't hang on macOS 26.
Phase 3 (read-back) uses openpyxl directly, because openpyxl's save
strips PHPP features (data validation extensions, headers/footers) that
make the file unopenable by Excel 2021 via AppleScript.
"""

from __future__ import annotations

import json
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from compare_json.compare import compare, format_report
from phpp_tool.reader import read_phpp
from phpp_tool.writer import write_phpp

FIELD_MAP = str(ROOT / "phpp-field-mapping.md")


def _read_back_openpyxl(
    written_path: Path,
    field_map_path: str,
) -> dict:
    """Read back a written workbook using openpyxl (no Excel needed).

    openpyxl's save strips features that make the file unopenable by
    Excel via AppleScript, so we use openpyxl for the read-back too.
    """
    import warnings
    warnings.filterwarnings("ignore", category=UserWarning)

    from openpyxl import load_workbook
    from phpp_tool.locators import col_to_idx, field_col, norm
    from phpp_tool.map_parser import parse_field_map

    wb = load_workbook(str(written_path), data_only=True)
    field_map = parse_field_map(field_map_path)
    sheet_names = wb.sheetnames
    result: dict = {}

    for ws_key, ws_spec in field_map.items():
        sheet_name = ws_spec["sheet_name"]
        si_name = f"{sheet_name} SI"
        if si_name in sheet_names:
            sheet_name = si_name
        if sheet_name not in sheet_names:
            continue

        ws = wb[sheet_name]
        ws_result: dict = {}

        for field_name, spec in ws_spec.get("fields", {}).items():
            loc_str = spec.get("locator_string")
            if not loc_str:
                continue
            loc_col = col_to_idx(spec["locator_col"])
            input_col = col_to_idx(spec["input_col"])
            offset = spec.get("row_offset", 0)
            needle = norm(loc_str)

            for row in range(1, ws.max_row + 1):
                cell_val = ws.cell(row=row, column=loc_col).value
                if cell_val and needle in norm(cell_val):
                    val = ws.cell(row=row + offset, column=input_col).value
                    if val is not None:
                        ws_result[field_name] = val
                    break

        if ws_result:
            result[ws_key] = ws_result

    wb.close()
    return result


def roundtrip(source: Path, template: Path, out_dir: Path) -> dict:
    source_name = source.stem
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n{'=' * 60}")
    print(f"  ROUNDTRIP: {source.name} → {template.name}")
    print(f"{'=' * 60}")

    # Step 1: read source workbook via xlwings (full fidelity)
    print(f"\n[1/4] Reading {source.name} (xlwings) ...")
    t0 = time.time()
    data1 = read_phpp(str(source), FIELD_MAP)
    t_read1 = time.time() - t0
    print(f"      {len(data1)} worksheet keys, {t_read1:.1f}s")

    json1_path = out_dir / f"{source_name}_read1.json"
    json1_path.write_text(
        json.dumps(data1, indent=2, default=str), encoding="utf-8"
    )

    # Step 2: write into template (xlwings locators + openpyxl persistence)
    written_path = out_dir / f"{source_name}_written.xlsx"
    print(f"\n[2/4] Writing into {template.name} → {written_path.name} ...")
    t0 = time.time()
    write_phpp(data1, str(template), str(written_path), FIELD_MAP)
    t_write = time.time() - t0
    print(f"      {t_write:.1f}s")

    # Step 3: read back via openpyxl (label-anchored fields only)
    print(f"\n[3/4] Reading back {written_path.name} (openpyxl) ...")
    t0 = time.time()
    data2 = _read_back_openpyxl(written_path, FIELD_MAP)
    t_read2 = time.time() - t0
    print(f"      {len(data2)} worksheet keys, {t_read2:.1f}s")

    json2_path = out_dir / f"{source_name}_read2.json"
    json2_path.write_text(
        json.dumps(data2, indent=2, default=str), encoding="utf-8"
    )

    # Step 4: compare (only the fields that _read_back_openpyxl can resolve)
    print(f"\n[4/4] Comparing label-anchored fields ...")
    data1_flat = {}
    for ws_key, ws_data in data1.items():
        if not isinstance(ws_data, dict):
            continue
        flat = {k: v for k, v in ws_data.items()
                if not isinstance(v, (dict, list))}
        if flat:
            data1_flat[ws_key] = flat

    results = compare(data1_flat, data2)
    report = format_report(results, str(json1_path.name), str(json2_path.name))

    report_path = out_dir / f"{source_name}_report.txt"
    report_path.write_text(report + "\n", encoding="utf-8")
    print(report)

    print(f"\n  Timing: read {t_read1:.1f}s + write {t_write:.1f}s"
          f" + read-back {t_read2:.1f}s"
          f" = {t_read1 + t_write + t_read2:.1f}s total")
    print(f"  Artifacts saved to {out_dir}/")

    return {
        "source": source.name,
        "template": template.name,
        "fields_read1": results["total_fields_file1"],
        "fields_read2": results["total_fields_file2"],
        "total_diffs": results["total_diffs"],
        "meaningful": results["meaningful"],
        "data_fidelity_pct": results["data_fidelity_pct"],
        "time_read1": t_read1,
        "time_write": t_write,
        "time_read2": t_read2,
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
            total_t = s["time_read1"] + s["time_write"] + s["time_read2"]
            print(f"  {s['source']:20s} → {s['template']:20s}"
                  f"  diffs: {s['total_diffs']:>4}"
                  f"  meaningful: {s['meaningful']:>3}"
                  f"  fidelity: {s['data_fidelity_pct']:.1f}%"
                  f"  time: {total_t:.1f}s")


if __name__ == "__main__":
    main()
