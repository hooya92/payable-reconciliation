from dataclasses import dataclass
from decimal import Decimal
from enum import Enum


class Status(str, Enum):
    MATCHED = "자동 대사"
    DESCRIPTION_MISMATCH = "적요 불일치"
    UNPAID = "지급 확인 안 됨"
    AMBIGUOUS = "중복/분할·합산 확인"
    VENDOR_MISMATCH = "거래처 오류 의심"


@dataclass(frozen=True)
class PayableItem:
    vendor_code: str
    vendor_name: str
    description: str
    amount: Decimal
    date: str = ""
    row_number: int = 0


@dataclass(frozen=True)
class JournalLine:
    vendor_code: str
    vendor_name: str
    account_code: str
    account_name: str
    description: str
    debit: Decimal
    credit: Decimal
    date: str = ""
    row_number: int = 0


@dataclass
class ReconcileResult:
    status: Status
    prior: PayableItem
    journal: JournalLine | None = None
    reason: str = ""
