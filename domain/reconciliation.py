from collections import defaultdict
from decimal import Decimal
from .models import JournalLine, NewPayable, PayableItem, ReconcileResult, Status
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
    """Text typo hint is limited to letters/digits; punctuation differences stay strict mismatches."""
    left=normalize_text(left); right=normalize_text(right)
    if left == right or abs(len(left)-len(right)) > 1:
        return False
    if len(left) == len(right):
        diff=[i for i,(a,b) in enumerate(zip(left,right)) if a != b]
        if len(diff) == 1:
            i=diff[0]
            return left[i].isalnum() and right[i].isalnum()
        return (
            len(diff) == 2
            and diff[1] == diff[0] + 1
            and left[diff[0]] == right[diff[1]]
            and left[diff[1]] == right[diff[0]]
            and all(left[i].isalnum() and right[i].isalnum() for i in diff)
        )

    short,long=(left,right) if len(left)<len(right) else (right,left)
    i=j=0
    while i<len(short) and short[i] == long[j]:
        i+=1; j+=1
    extra=long[j] if j<len(long) else long[-1]
    if not extra.isalnum():
        return False
    return short[i:] == long[j+1:]


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
    prior_desc_counts: dict[tuple[str, str], int] = defaultdict(int)
    for item in prior_items:
        prior_key_counts[(normalize_code(item.vendor_code), item.amount)] += 1
        prior_desc_counts[(normalize_code(item.vendor_code), normalize_text(item.description))] += 1

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
        if len(candidates) == 1:
            line=candidates[0]
            # A same-amount debit might settle a different opening item partially.
            competing_prior = any(
                other is not item
                and normalize_code(other.vendor_code) == key[0]
                and other.amount > line.debit
                and normalize_text(other.description) == normalize_text(line.description)
                for other in prior_items
            )
            if competing_prior and normalize_text(item.description) != normalize_text(line.description):
                results.append(ReconcileResult(
                    Status.AMBIGUOUS, item, line,
                    "동일 차변이 다른 전월 항목의 부분지급일 수 있어 자동 배정하지 않음",
                    "COMPETING_PARTIAL"
                ))
                continue
            used.add(id(line))

            same_name=normalize_text(line.vendor_name) == normalize_text(item.vendor_name)
            same_desc=normalize_text(line.description) == normalize_text(item.description)
            notes=[]
            if not same_name:
                notes.append("거래처명 오타 의심" if _text_typo(line.vendor_name,item.vendor_name) else "거래처명 차이")
            if not same_desc:
                notes.append("적요 오타 의심" if _text_typo(line.description,item.description) else "적요 차이")

            if notes:
                results.append(ReconcileResult(
                    Status.MATCHED, item, line,
                    "거래처코드·금액이 유일하게 일치 · 참고: " + " / ".join(notes),
                    "CODE_AMOUNT_UNIQUE_WITH_NOTE"
                ))
            else:
                results.append(ReconcileResult(
                    Status.MATCHED, item, line,
                    "거래처코드·금액·적요 일치", "EXACT"
                ))

        elif len(exact) > 1 or len(candidates) > 1:
            results.append(ReconcileResult(
                Status.AMBIGUOUS, item, (exact or candidates)[0],
                "동일 거래처/금액 후보가 여러 건", "DUPLICATE"
            ))

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
            exact_partial = [
                x for x in partial
                if normalize_text(x.description) == normalize_text(item.description)
            ]
            if (
                len(exact_partial) == 1
                and len(remaining_vendor_debits) == 1
                and prior_desc_counts[(key[0], normalize_text(item.description))] == 1
                and not any(
                    other is not item
                    and normalize_code(other.vendor_code) == key[0]
                    and other.amount == exact_partial[0].debit
                    for other in prior_items
                )
            ):
                line = exact_partial[0]
                used.add(id(line))
                balance = item.amount - line.debit
                results.append(ReconcileResult(
                    Status.PARTIAL, item, line,
                    f"전월 {item.amount:,.0f}원 중 당월 {line.debit:,.0f}원 지급 · 잔액 {balance:,.0f}원 이월",
                    "PARTIAL_UNIQUE"
                ))
                continue
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


def new_payables(
    journal_lines: list[JournalLine],
    prior_results: list[ReconcileResult],
    allow_auto: bool = True,
) -> list[NewPayable]:
    """Only carry current credits whose same-month settlement is unambiguous."""
    credits = [line for line in journal_lines if line.credit > 0]
    used_debits = {
        id(r.journal) for r in prior_results
        if r.status in (Status.MATCHED, Status.PARTIAL) and r.journal is not None
    }
    spare_debits = [
        line for line in journal_lines
        if line.debit > 0 and id(line) not in used_debits
    ]
    credit_counts = defaultdict(int)
    prior_keys = {
        (normalize_code(r.prior.vendor_code), normalize_text(r.prior.description))
        for r in prior_results
    }
    for line in credits:
        credit_counts[(normalize_code(line.vendor_code), normalize_text(line.description))] += 1

    new_items = []
    for line in credits:
        code = normalize_code(line.vendor_code)
        key = (code, normalize_text(line.description))
        vendor_debits = [d for d in spare_debits if normalize_code(d.vendor_code) == code]
        same_desc_debits = [d for d in vendor_debits if normalize_text(d.description) == key[1]]
        reason = ""
        if not allow_auto:
            reason = "입력 데이터 확인 항목이 있어 신규 발생분 자동 이월 보류"
        elif credit_counts[key] != 1 or key in prior_keys:
            reason = "전월 항목과 중복 또는 당월 동일 적요 대변 중복 가능성 확인 필요"
        elif len(same_desc_debits) > 1 or len(vendor_debits) != len(same_desc_debits):
            reason = "당월 차변을 신규 대변에 안전하게 배정할 수 없어 확인 필요"
        elif sum(d.debit for d in same_desc_debits) > line.credit:
            reason = "당월 차변 합계가 신규 대변보다 커서 확인 필요"

        if reason:
            new_items.append(NewPayable(line, line.credit, reason, False))
        else:
            remaining = line.credit - sum(d.debit for d in same_desc_debits)
            description = (
                "당월 차변과 전액 상계"
                if remaining == 0 else
                "당월 차변 일부 상계 후 잔액 이월"
                if same_desc_debits else
                "당월 신규 발생 미지급금 이월"
            )
            new_items.append(NewPayable(line, remaining, description, True))
    return new_items
