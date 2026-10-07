from dataclasses import dataclass
from datetime import date, datetime


@dataclass(frozen=True, order=True)
class AccountingPeriod:
    year: int
    month: int

    def __post_init__(self):
        if not 1 <= self.month <= 12:
            raise ValueError("월은 1~12 사이여야 합니다.")

    def previous(self) -> "AccountingPeriod":
        return AccountingPeriod(self.year - 1, 12) if self.month == 1 else AccountingPeriod(self.year, self.month - 1)

    def next(self) -> "AccountingPeriod":
        return AccountingPeriod(self.year + 1, 1) if self.month == 12 else AccountingPeriod(self.year, self.month + 1)

    @property
    def label(self) -> str:
        return f"{self.year}년 {self.month}월"

    def contains(self, value: object) -> bool:
        parsed = parse_date(value)
        return parsed is not None and parsed.year == self.year and parsed.month == self.month


def parse_date(value: object) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if value is None:
        return None
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%Y%m%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def prior_period_for(current_year: int, current_month: int) -> AccountingPeriod:
    return AccountingPeriod(current_year, current_month).previous()
