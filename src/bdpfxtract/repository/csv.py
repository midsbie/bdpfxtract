from __future__ import annotations

import csv
from decimal import Decimal
from pathlib import Path

from bdpfxtract.logging.logging import configure_logging
from bdpfxtract.model import ForexRecord

logger = configure_logging()


class CsvRepository:
    """Loads and persists ForexRecord items to CSV with merge semantics."""

    header = ("Date", "Currency", "Price")

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
                k = (row["Date"], row["Currency"])  # type: ignore[index]
                try:
                    data[k] = Decimal(row["Price"])  # type: ignore[index]
                except Exception as exc:  # noqa: BLE001
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
                        "Date": date_str,
                        "Currency": currency,
                        "Price": str(price),
                    }
                )

    @staticmethod
    def key_for(r: ForexRecord) -> tuple[str, str]:
        return (r.date.isoformat(), r.currency)
