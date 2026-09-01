#!/usr/bin/env python3
"""Builds the Canadian Finance Tracker workbooks.

    python3 build.py                      # dist/*.xlsm, dist/*.xlsx and samples/
    python3 build.py --today 2026-09-01   # reproducible build for a fixed date

Two editions come out of one source: the macro-enabled workbook (.xlsm), which
imports bank exports, and the plain workbook (.xlsx), which has the same sheets
and reports but is filled in by hand and needs no macros.

The sample transactions are dated relative to "today" so that a freshly built
workbook always opens on a month that has data in it.  Pass --today to get
byte-for-byte reproducible artifacts.
"""

from __future__ import annotations

import argparse
import hashlib
from datetime import date
from pathlib import Path

from tools import package, sample, vbaproject, workbook

ROOT = Path(__file__).resolve().parent
DIST = ROOT / "dist"
SAMPLES = ROOT / "samples"
ARTIFACT = "Canadian-Finance-Tracker.xlsm"
PLAIN_ARTIFACT = "Canadian-Finance-Tracker.xlsx"


def parse_day(text: str) -> date:
    return date.fromisoformat(text)


def build_package(today: date) -> bytes:
    """The macro-enabled edition."""
    wb = workbook.build(today, macros=True)
    code_names = [wb[name].sheet_properties.codeName for name in wb.sheetnames]
    vba = vbaproject.build(code_names)
    return package.to_xlsm(wb, vba)


def build_plain_package(today: date) -> bytes:
    """The edition without macros."""
    return package.to_xlsx(workbook.build(today, macros=False))


def write_samples(today: date, directory: Path = SAMPLES) -> list[Path]:
    directory.mkdir(parents=True, exist_ok=True)
    written = []
    for name, text in sample.csv_files(sample.build(today), today).items():
        path = directory / name
        path.write_text(text, encoding="utf-8", newline="")
        written.append(path)
    return written


def _report(path: Path, blob: bytes) -> None:
    digest = hashlib.sha256(blob).hexdigest()[:16]
    shown = path.relative_to(ROOT) if path.is_relative_to(ROOT) else path
    print(f"{shown}  {len(blob):,} bytes  sha256:{digest}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--today", type=parse_day, default=date.today(),
                        help="date the sample data is generated relative to")
    parser.add_argument("--dist", type=Path, default=DIST,
                        help="directory the two workbooks are written to")
    parser.add_argument("--no-samples", action="store_true")
    args = parser.parse_args()

    args.dist.mkdir(parents=True, exist_ok=True)
    for name, blob in ((ARTIFACT, build_package(args.today)),
                       (PLAIN_ARTIFACT, build_plain_package(args.today))):
        path = args.dist / name
        path.write_bytes(blob)
        _report(path, blob)

    if not args.no_samples:
        for path in write_samples(args.today):
            print(f"{path.relative_to(ROOT)}  {path.stat().st_size:,} bytes")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
