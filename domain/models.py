from dataclasses import dataclass, field
from decimal import Decimal
from enum import Enum


class Status(str, Enum):
    MATCHED = "자동 대사"
    DESCRIPTION_MISMATCH = "적요 불일치"
    VENDOR_NAME_MISMATCH = "거래처명 변경 확인"
    UNPAID = "지급 확인 안 됨"
    AMBIGUOUS = "중복/분할·합산 확인"
    VENDOR_MISMATCH = "거래처 오류 의심"
    INVALID_INPUT = "입력 데이터 확인"


class DataQuality(str, Enum):
    CLEAN = "정상"
    NORMALIZED = "형식 보정"
    SUSPICIOUS = "확인 필요"


@dataclass(frozen=True)
class NormalizedAmount:
    raw: object
    value: Decimal | None
    quality: DataQuality
    reason: str = ""


@dataclass(frozen=True)
class SourceRef:
    file_name: str = ""
    sheet: str = ""
    row: int = 0
    owner: str = ""


@dataclass(frozen=True)
class PayableItem:
    vendor_code: str
    vendor_name: str
    description: str
    amount: Decimal
    date: str = ""
    row_number: int = 0
    source: SourceRef = field(default_factory=SourceRef)


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
    rule: str = ""
    source_rows: list[int] = field(default_factory=list)
