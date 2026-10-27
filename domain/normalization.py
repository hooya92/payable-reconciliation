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

def parse_literal_amount_formula(value: object) -> Decimal | None:
    """Read only a formula containing a signed numeric constant, never references."""
    if not isinstance(value,str):
        return None
    match=re.fullmatch(
        r"=\s*([+-]?)\s*(?:(\d+(?:\.\d+)?)|\(\s*([+-]?)\s*(\d+(?:\.\d+)?)\s*\))\s*",value)
    if not match:
        return None
    amount=Decimal(match.group(2) or match.group(4))
    return -amount if (match.group(1)=="-") != (match.group(3)=="-") else amount


def parse_amount(value: object, *, allow_zero_marker: bool = True) -> NormalizedAmount:
    """Parse each amount independently; reject ambiguous literal dash placeholders in statements.

    Journal debit/credit columns may use '-' as an empty/zero-side marker, but
    a statement's amount cell must contain an actual numeric value to be verified.
    """
    if value is None or value == "":
        return NormalizedAmount(value, Decimal("0"), DataQuality.CLEAN)
    if isinstance(value, (int, Decimal)):
        return NormalizedAmount(value, Decimal(value), DataQuality.CLEAN)
    if isinstance(value, float):
        return NormalizedAmount(value, Decimal(str(value)), DataQuality.CLEAN)

    raw = str(value).strip()
    # Fold full-width amount characters only; NFKC would also turn 10² into 102.
    normalized=raw.translate(str.maketrans("０１２３４５６７８９，．（）＋－￦","0123456789,.()+-₩")).translate(
        str.maketrans({char:"-" for char in "−﹣－‒–—‐‑"})
    )
    s=normalized
    prefix=s.startswith("₩")
    suffix=s.endswith("원")
    if prefix:
        s=s[1:].strip()
    if suffix:
        s=s[:-1].strip()
    if s=="-":
        if not allow_zero_marker:
            return NormalizedAmount(value,None,DataQuality.SUSPICIOUS,
                                    "숫자 없는 마이너스 기호는 0원으로 확정할 수 없음")
        return NormalizedAmount(value,Decimal(0),DataQuality.NORMALIZED,"회계식 0 표시")
    parentheses=s.startswith("(") and s.endswith(")")
    if parentheses:
        s=s[1:-1].strip()
    match=re.fullmatch(
        r"([+-]?)\s*(₩?)\s*(\d+|\d{1,3}(?:,\d{3})+)\s*(원?)\s*(-?)",s)
    if match:
        sign,currency,number,unit,trailing=match.groups()
        # Never guess the meaning of conflicting signs or repeated currency units.
        if (not (parentheses and (sign or trailing)) and not (sign and trailing)
                and not (prefix and currency) and not (suffix and unit)):
            amount=Decimal(number.replace(",",""))
            if sign=="-" or parentheses or trailing:
                amount=-amount
            changed=prefix or suffix or currency or unit or parentheses or trailing
            return NormalizedAmount(value,amount,
                DataQuality.NORMALIZED if changed else DataQuality.CLEAN,
                "회계식 금액 표기 보정" if changed else "")
    # Allow spacing between a leading sign and the amount, including NBSP
    # from Excel copy/paste. Spaces inside digits must never join two numbers.
    s = re.sub(r"^([+\-])\s*",r"\1",normalized)
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
