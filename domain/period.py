from dataclasses import dataclass


@dataclass(frozen=True, order=True)
class AccountingPeriod:
    year: int
    month: int

    def __post_init__(self):
        if not 1 <= self.month <= 12:
            raise ValueError("월은 1~12 사이여야 합니다.")

    def previous(self) -> "AccountingPeriod":
        return AccountingPeriod(self.year - 1, 12) if self.month == 1 else AccountingPeriod(self.year, self.month - 1)

    @property
    def label(self) -> str:
        return f"{self.year}년 {self.month}월"


def prior_period_for(current_year: int, current_month: int) -> AccountingPeriod:
    return AccountingPeriod(current_year, current_month).previous()
