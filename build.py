#!/usr/bin/env python3
"""Builds the Canadian Finance Tracker workbook.

    python3 build.py                 # dist/Canadian-Finance-Tracker.xlsm + samples/
    python3 build.py --today 2026-09-01   # reproducible build for a fixed date

The sample transactions are dated relative to "today" so that a freshly built
workbook always opens on a month that has data in it.  Pass --today to get a
byte-for-byte reproducible artifact.
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


def parse_day(text: str) -> date:
    return date.fromisoformat(text)


def build_package(today: date) -> bytes:
    wb = workbook.build(today)
    code_names = [wb[name].sheet_properties.codeName for name in wb.sheetnames]
    vba = vbaproject.build(code_names)
    return package.to_xlsm(wb, vba)


def write_samples(today: date, directory: Path = SAMPLES) -> list[Path]:
    directory.mkdir(parents=True, exist_ok=True)
    written = []
    for name, text in sample.csv_files(sample.build(today), today).items():
        path = directory / name
        path.write_text(text, encoding="utf-8", newline="")
        written.append(path)
    return written


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--today", type=parse_day, default=date.today(),
                        help="date the sample data is generated relative to")
    parser.add_argument("--out", type=Path, default=DIST / ARTIFACT)
    parser.add_argument("--no-samples", action="store_true")
    args = parser.parse_args()

    blob = build_package(args.today)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_bytes(blob)

    digest = hashlib.sha256(blob).hexdigest()[:16]
    print(f"{args.out.relative_to(ROOT) if args.out.is_relative_to(ROOT) else args.out}"
          f"  {len(blob):,} bytes  sha256:{digest}")

    if not args.no_samples:
        for path in write_samples(args.today):
            print(f"{path.relative_to(ROOT)}  {path.stat().st_size:,} bytes")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
