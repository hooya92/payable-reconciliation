import re
from decimal import Decimal, InvalidOperation
from .models import DataQuality, NormalizedAmount


def normalize_code(value: object) -> str:
    """Codes are identifiers, never dates. Do not infer or pad missing leading zeroes."""
    if value is None:
        return ""
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return re.sub(r"\s+", "", text)


def normalize_text(value: object) -> str:
    """Only normalize harmless whitespace/case differences; preserve punctuation and wording."""
    text = "" if value is None else str(value)
    return re.sub(r"\s+", " ", text.strip().casefold())


def parse_amount(value: object) -> NormalizedAmount:
    """Parse only deterministic formats. Suspicious separators/units are never auto-approved."""
    if value is None or value == "":
        return NormalizedAmount(value, Decimal("0"), DataQuality.CLEAN)
    if isinstance(value, (int, Decimal)):
        return NormalizedAmount(value, Decimal(value), DataQuality.CLEAN)
    if isinstance(value, float):
        return NormalizedAmount(value, Decimal(str(value)), DataQuality.CLEAN)

    raw = str(value).strip()
    s = raw.replace(" ", "")
    if re.fullmatch(r"[+-]?\d+", s):
        return NormalizedAmount(value, Decimal(s), DataQuality.CLEAN)
    if re.fullmatch(r"[+-]?\d{1,3}(,\d{3})+", s):
        return NormalizedAmount(value, Decimal(s.replace(",", "")), DataQuality.CLEAN)
    if re.fullmatch(r"[+-]?(?:₩)?\d{1,3}(?:,\d{3})*(?:원)?", s) and ("₩" in s or "원" in s):
        cleaned = s.replace("₩", "").replace("원", "").replace(",", "")
        return NormalizedAmount(value, Decimal(cleaned), DataQuality.NORMALIZED, "통화 기호/단위 제거")
    if re.fullmatch(r"[+-]?\d{1,3}(\.\d{3})+", s):
        guessed = Decimal(s.replace(".", ""))
        return NormalizedAmount(value, guessed, DataQuality.SUSPICIOUS, "점(.) 천단위 구분자는 확인 필요")
    if "만원" in s:
        m = re.fullmatch(r"([+-]?\d+(?:\.\d+)?)만원", s)
        guessed = Decimal(m.group(1)) * 10000 if m else None
        return NormalizedAmount(value, guessed, DataQuality.SUSPICIOUS, "만원 단위 표기는 자동 대사하지 않음")
    try:
        guessed = Decimal(re.sub(r"[^0-9+\-]", "", s))
    except (InvalidOperation, ValueError):
        guessed = None
    return NormalizedAmount(value, guessed, DataQuality.SUSPICIOUS, "비표준 금액 형식")
