import re
from collections import defaultdict
from decimal import Decimal
from models import JournalLine, PayableItem, ReconcileResult, Status


def normalize_code(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return re.sub(r"\s+", "", text)


def normalize_text(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.lower().strip()
    return re.sub(r"[\s_\-./·,()\[\]]+", "", text)


def reconcile(prior_items: list[PayableItem], journal_lines: list[JournalLine]) -> list[ReconcileResult]:
    """Conservative reconciliation: only exact vendor+amount+normalized-description is auto matched."""
    debits = [j for j in journal_lines if j.debit > 0]
    by_key: dict[tuple[str, Decimal], list[JournalLine]] = defaultdict(list)
    by_amount: dict[Decimal, list[JournalLine]] = defaultdict(list)
    for line in debits:
        by_key[(normalize_code(line.vendor_code), line.debit)].append(line)
        by_amount[line.debit].append(line)

    used: set[int] = set()
    results: list[ReconcileResult] = []

    for item in prior_items:
        key = (normalize_code(item.vendor_code), item.amount)
        candidates = [x for x in by_key.get(key, []) if id(x) not in used]
        exact = [x for x in candidates if normalize_text(x.description) == normalize_text(item.description)]

        if len(exact) == 1:
            used.add(id(exact[0]))
            results.append(ReconcileResult(Status.MATCHED, item, exact[0], "거래처코드·금액·적요 일치"))
        elif len(exact) > 1 or len(candidates) > 1:
            results.append(ReconcileResult(Status.AMBIGUOUS, item, exact[0] if exact else candidates[0],
                                           "동일 거래처/금액 후보가 여러 건"))
        elif len(candidates) == 1:
            results.append(ReconcileResult(Status.DESCRIPTION_MISMATCH, item, candidates[0],
                                           "거래처코드와 금액은 같지만 적요가 다름"))
        else:
            other_vendor = [x for x in by_amount.get(item.amount, [])
                            if normalize_code(x.vendor_code) != normalize_code(item.vendor_code)]
            if other_vendor:
                results.append(ReconcileResult(Status.VENDOR_MISMATCH, item, other_vendor[0],
                                               "다른 거래처코드에서 동일 차변 금액 발견"))
            else:
                results.append(ReconcileResult(Status.UNPAID, item, None, "대응하는 차변을 찾지 못함"))
    return results


def new_payables(journal_lines: list[JournalLine]) -> list[JournalLine]:
    return [line for line in journal_lines if line.credit > 0]
