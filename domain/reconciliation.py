from collections import defaultdict
from decimal import Decimal
from .models import JournalLine, PayableItem, ReconcileResult, Status
from .normalization import normalize_code, normalize_text


def _single_edit_or_transposition(left: str, right: str) -> bool:
    """True only for a very small manual-entry difference; never used for auto matching."""
    if left == right:
        return False
    if abs(len(left) - len(right)) > 1:
        return False
    if len(left) == len(right):
        diff=[i for i,(a,b) in enumerate(zip(left,right)) if a != b]
        if len(diff) == 1:
            return True
        return (
            len(diff) == 2
            and diff[1] == diff[0] + 1
            and left[diff[0]] == right[diff[1]]
            and left[diff[1]] == right[diff[0]]
        )
    short,long=(left,right) if len(left)<len(right) else (right,left)
    i=j=0; skipped=False
    while i<len(short) and j<len(long):
        if short[i] == long[j]:
            i+=1; j+=1
        elif skipped:
            return False
        else:
            skipped=True; j+=1
    return True


def _text_typo(left: object, right: object) -> bool:
    return _single_edit_or_transposition(normalize_text(left), normalize_text(right))


def _code_typo(left: object, right: object) -> bool:
    return _single_edit_or_transposition(normalize_code(left), normalize_code(right))


def _amount_typo(left: Decimal, right: Decimal) -> bool:
    if left <= 0 or right <= 0 or left != left.to_integral_value() or right != right.to_integral_value():
        return False
    return _single_edit_or_transposition(str(int(left)), str(int(right)))


def reconcile(prior_items: list[PayableItem], journal_lines: list[JournalLine]) -> list[ReconcileResult]:
    """Dates never decide whether a payable matches. Period is only an input-file scope."""
    debits = [j for j in journal_lines if j.debit > 0]
    by_key: dict[tuple[str, Decimal], list[JournalLine]] = defaultdict(list)
    by_amount: dict[Decimal, list[JournalLine]] = defaultdict(list)
    by_vendor: dict[str, list[JournalLine]] = defaultdict(list)
    for line in debits:
        by_key[(normalize_code(line.vendor_code), line.debit)].append(line)
        by_amount[line.debit].append(line)
        by_vendor[normalize_code(line.vendor_code)].append(line)

    used: set[int] = set()
    results: list[ReconcileResult] = []
    prior_key_counts: dict[tuple[str, Decimal], int] = defaultdict(int)
    for item in prior_items:
        prior_key_counts[(normalize_code(item.vendor_code), item.amount)] += 1

    for item in prior_items:
        key = (normalize_code(item.vendor_code), item.amount)
        candidates = [x for x in by_key.get(key, []) if id(x) not in used]

        if prior_key_counts[key] > 1 and len(candidates) < prior_key_counts[key]:
            results.append(ReconcileResult(
                Status.AMBIGUOUS, item, candidates[0] if candidates else None,
                "전월 명세에 동일 거래처/금액이 중복되어 1:1 할당 불가", "PRIOR_DUPLICATE"
            ))
            continue

        exact = [x for x in candidates if normalize_text(x.description) == normalize_text(item.description)]
        if len(exact) == 1 and len(candidates) == 1:
            line=exact[0]
            used.add(id(line))
            if normalize_text(line.vendor_name) != normalize_text(item.vendor_name):
                if _text_typo(line.vendor_name,item.vendor_name):
                    results.append(ReconcileResult(
                        Status.INPUT_TYPO_SUSPECT, item, line,
                        "거래처코드·금액·적요는 일치하고 거래처명만 1글자 수준 차이 — 거래처명 수기 오타 의심",
                        "VENDOR_NAME_TYPO"
                    ))
                else:
                    results.append(ReconcileResult(
                        Status.VENDOR_NAME_MISMATCH, item, line,
                        "거래처코드는 같지만 거래처명이 다름", "VENDOR_NAME"
                    ))
            else:
                results.append(ReconcileResult(Status.MATCHED, item, line, "거래처코드·금액·적요 일치", "EXACT"))

        elif len(exact) > 1 or len(candidates) > 1:
            results.append(ReconcileResult(
                Status.AMBIGUOUS, item, (exact or candidates)[0],
                "동일 거래처/금액 후보가 여러 건", "DUPLICATE"
            ))

        elif len(candidates) == 1:
            line=candidates[0]
            used.add(id(line))
            same_name=normalize_text(line.vendor_name) == normalize_text(item.vendor_name)
            if same_name and _text_typo(line.description,item.description):
                results.append(ReconcileResult(
                    Status.INPUT_TYPO_SUSPECT, item, line,
                    "거래처코드·거래처명·금액은 일치하고 적요만 1글자 수준 차이 — 적요 수기 오타 의심",
                    "DESCRIPTION_TYPO"
                ))
            else:
                reason="거래처코드와 금액은 같지만 적요가 다름"
                if not same_name:
                    reason += " · 거래처명도 다름"
                results.append(ReconcileResult(Status.DESCRIPTION_MISMATCH, item, line, reason, "VENDOR_AMOUNT"))

        else:
            remaining_vendor_debits = [
                x for x in by_vendor.get(key[0], [])
                if id(x) not in used and x.debit > 0
            ]

            amount_typo = [
                x for x in remaining_vendor_debits
                if normalize_text(x.vendor_name) == normalize_text(item.vendor_name)
                and normalize_text(x.description) == normalize_text(item.description)
                and _amount_typo(x.debit,item.amount)
            ]
            if len(amount_typo) == 1:
                line=amount_typo[0]
                results.append(ReconcileResult(
                    Status.INPUT_TYPO_SUSPECT, item, line,
                    f"거래처코드·거래처명·적요는 일치하지만 금액이 {item.amount:,.0f}원 ↔ {line.debit:,.0f}원 — 금액 수기 오타 또는 실제 부분/합산 지급 여부 확인",
                    "AMOUNT_TYPO"
                ))
                continue
            if len(amount_typo) > 1:
                results.append(ReconcileResult(
                    Status.AMBIGUOUS, item, amount_typo[0],
                    "거래처·적요는 일치하지만 금액이 비슷한 후보가 여러 건 — 금액 오타/분할·합산 여부 확인",
                    "AMOUNT_TYPO_MULTIPLE"
                ))
                continue

            partial = [x for x in remaining_vendor_debits if x.debit < item.amount]
            if partial:
                results.append(ReconcileResult(
                    Status.AMBIGUOUS, item, partial[0],
                    "동일 거래처에 더 작은 차변이 있어 부분지급 가능성 확인 필요", "POSSIBLE_PARTIAL"
                ))
                continue

            combined = [x for x in remaining_vendor_debits if x.debit > item.amount]
            if combined:
                results.append(ReconcileResult(
                    Status.AMBIGUOUS, item, combined[0],
                    "동일 거래처에 더 큰 차변이 있어 합산지급 가능성 확인 필요", "POSSIBLE_COMBINED"
                ))
                continue

            other = [
                x for x in by_amount.get(item.amount, [])
                if normalize_code(x.vendor_code) != normalize_code(item.vendor_code)
            ]
            code_typo = [
                x for x in other
                if normalize_text(x.vendor_name) == normalize_text(item.vendor_name)
                and normalize_text(x.description) == normalize_text(item.description)
                and _code_typo(x.vendor_code,item.vendor_code)
            ]
            if len(code_typo) == 1:
                line=code_typo[0]
                results.append(ReconcileResult(
                    Status.INPUT_TYPO_SUSPECT, item, line,
                    f"거래처명·금액·적요는 일치하지만 거래처코드가 {item.vendor_code} ↔ {line.vendor_code} — 거래처코드 수기 오타 의심",
                    "VENDOR_CODE_TYPO"
                ))
            elif other:
                results.append(ReconcileResult(
                    Status.VENDOR_MISMATCH, item, other[0],
                    "다른 거래처코드에서 동일 차변 금액 발견", "OTHER_VENDOR_AMOUNT"
                ))
            else:
                results.append(ReconcileResult(Status.UNPAID, item, None, "대응하는 차변을 찾지 못함", "NO_DEBIT"))
    return results


def new_payables(journal_lines: list[JournalLine]) -> list[JournalLine]:
    return [line for line in journal_lines if line.credit > 0]
