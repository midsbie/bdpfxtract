from __future__ import annotations

import argparse
import datetime as dt
from pathlib import Path
from typing import Iterable, Optional

from bdpfxtract.logging import configure_logging
from bdpfxtract.repository import CsvRepository
from bdpfxtract.parser import BdpForexParser

logger = configure_logging()


class BdpForexToCsvApp:
    """Application entry point to parse an XLSX and merge/write a CSV.

    Usage:
      bdpfxtract sample.xlsx out.csv
      python -m bdpfxtract.cmd.cli sample.xlsx out.csv
    """

    def __init__(self) -> None:
        self._parser = BdpForexParser()
        self._repo = CsvRepository()

    def run(self, xlsx_path: Path, out_csv: Path) -> Path:
        logger.info("Parsing workbook: %s", xlsx_path)
        records = self._parser.parse(xlsx_path)
        logger.info("Merging into CSV: %s", out_csv)

        existing = self._repo.load(out_csv)
        for r in records:
            existing[self._repo.key_for(r)] = r.price
        # Persist
        out_csv.parent.mkdir(parents=True, exist_ok=True)
        self._repo.save(out_csv, existing)
        logger.info("Wrote %d rows to %s", len(existing), out_csv)
        return out_csv


def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="bdpfxtract",
        description=(
            "Parse a Banco de Portugal FOREX XLSX and merge or create a CSV "
            "with columns: Date, Currency, Rate (FC per 1 EUR)."
        ),
    )
    p.add_argument("input", help="Path to XLSX input file")
    p.add_argument("output", help="Path to CSV output file")
    return p


def main(argv: Optional[Iterable[str]] = None) -> int:
    args = build_arg_parser().parse_args(argv)
    xlsx = Path(args.input)
    out_csv = Path(args.output)

    if not xlsx.exists():
        logger.error("Input XLSX not found: %s", xlsx)
        return 2

    try:
        BdpForexToCsvApp().run(xlsx, out_csv)
    except Exception as exc:
        logger.exception("Failed to process workbook: %s", exc)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
