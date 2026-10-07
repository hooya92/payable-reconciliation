from collections import defaultdict
from decimal import Decimal
from .models import JournalLine, PayableItem, ReconcileResult, Status
from .normalization import normalize_code, normalize_text


def reconcile(prior_items: list[PayableItem], journal_lines: list[JournalLine]) -> list[ReconcileResult]:
    """Dates never decide whether a payable matches. Period is only an input-file scope."""
    debits = [j for j in journal_lines if j.debit > 0]
    by_key: dict[tuple[str, Decimal], list[JournalLine]] = defaultdict(list)
    by_amount: dict[Decimal, list[JournalLine]] = defaultdict(list)
    for line in debits:
        by_key[(normalize_code(line.vendor_code), line.debit)].append(line)
        by_amount[line.debit].append(line)

    used: set[int] = set()
    results: list[ReconcileResult] = []
    prior_key_counts: dict[tuple[str, Decimal], int] = defaultdict(int)
    for item in prior_items:
        prior_key_counts[(normalize_code(item.vendor_code), item.amount)] += 1
    for item in prior_items:
        key = (normalize_code(item.vendor_code), item.amount)
        candidates = [x for x in by_key.get(key, []) if id(x) not in used]
        if prior_key_counts[key] > 1 and len(candidates) < prior_key_counts[key]:
            results.append(ReconcileResult(Status.AMBIGUOUS, item, candidates[0] if candidates else None, "전월 명세에 동일 거래처/금액이 중복되어 1:1 할당 불가", "PRIOR_DUPLICATE"))
            continue
        exact = [x for x in candidates if normalize_text(x.description) == normalize_text(item.description)]
        if len(exact) == 1:
            used.add(id(exact[0]))
            results.append(ReconcileResult(Status.MATCHED, item, exact[0], "거래처코드·금액·적요 일치", "EXACT"))
        elif len(exact) > 1 or len(candidates) > 1:
            results.append(ReconcileResult(Status.AMBIGUOUS, item, (exact or candidates)[0], "동일 거래처/금액 후보가 여러 건", "DUPLICATE"))
        elif len(candidates) == 1:
            used.add(id(candidates[0]))
            results.append(ReconcileResult(Status.DESCRIPTION_MISMATCH, item, candidates[0], "거래처코드와 금액은 같지만 적요가 다름", "VENDOR_AMOUNT"))
        else:
            other = [x for x in by_amount.get(item.amount, []) if normalize_code(x.vendor_code) != normalize_code(item.vendor_code)]
            if other:
                results.append(ReconcileResult(Status.VENDOR_MISMATCH, item, other[0], "다른 거래처코드에서 동일 차변 금액 발견", "OTHER_VENDOR_AMOUNT"))
            else:
                results.append(ReconcileResult(Status.UNPAID, item, None, "대응하는 차변을 찾지 못함", "NO_DEBIT"))
    return results


def new_payables(journal_lines: list[JournalLine]) -> list[JournalLine]:
    return [line for line in journal_lines if line.credit > 0]
