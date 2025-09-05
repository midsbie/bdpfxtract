import datetime as dt
from dataclasses import dataclass
from decimal import Decimal


@dataclass(frozen=True)
class ForexRecord:
    date: dt.date
    currency: str
    price: Decimal
