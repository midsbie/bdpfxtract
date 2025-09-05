from __future__ import annotations

import csv
from decimal import Decimal
from pathlib import Path

from bdpfxtract.logging import configure_logging
from bdpfxtract.model import ForexRecord

logger = configure_logging()


class CsvRepository:
    """Loads and persists ForexRecord items to CSV with merge semantics."""

    header = ("date", "currency", "rate")

    def load(self, path: Path) -> dict[tuple[str, str], Decimal]:
        data: dict[tuple[str, str], Decimal] = {}
        if not path.exists():
            return data
        with path.open("r", newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            missing_cols = [
                c
                for c in self.header
                if c not in reader.fieldnames or reader.fieldnames is None
            ]
            if missing_cols:
                raise ValueError(f"CSV missing required columns: {missing_cols}")
            for row in reader:
                k = (row["date"], row["currency"])  # type: ignore[index]
                try:
                    data[k] = Decimal(row["rate"])  # type: ignore[index]
                except Exception as exc:
                    logger.warning("Skipping invalid CSV row %r: %s", row, exc)
        return data

    def save(self, path: Path, kv: dict[tuple[str, str], Decimal]) -> None:
        # Sort for stable output: by date then currency
        rows = sorted(kv.items(), key=lambda x: (x[0][0], x[0][1]))
        with path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=list(self.header))
            writer.writeheader()
            for (date_str, currency), price in rows:
                writer.writerow(
                    {
                        "date": date_str,
                        "currency": currency,
                        "rate": str(price),
                    }
                )

    @staticmethod
    def key_for(r: ForexRecord) -> tuple[str, str]:
        return (r.date.isoformat(), r.currency)
