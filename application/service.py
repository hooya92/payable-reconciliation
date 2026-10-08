from collections import Counter
from dataclasses import dataclass
from pathlib import Path

from adapters.excel.reader import read_douzone, read_prior
from domain.models import ReconcileResult, Status
from domain.normalization import normalize_code, normalize_text
from domain.period import AccountingPeriod
from domain.reconciliation import new_payables, reconcile


@dataclass
class ReconciliationRun:
    period: AccountingPeriod
    prior_count: int
    results: list
    new_items: list
    issues: list
    counts: Counter


def run_reconciliation(prior_paths, douzone_paths, account_codes, period):
    if not account_codes:
        raise ValueError("미지급금 계정코드를 1개 이상 지정해야 합니다.")
    prior_items=[]; opening_reviews=[]; prior_issues=[]
    for path in prior_paths:
        rr=read_prior(path,Path(path).stem,period.previous())
        prior_items.extend(rr.items); opening_reviews.extend(rr.review_items); prior_issues.extend(rr.issues)

    review_vendor_seen={}
    for item,_reason in opening_reviews:
        code=normalize_code(item.vendor_code)
        if code in review_vendor_seen:
            previous=review_vendor_seen[code]
            raise ValueError(
                f"전월 음수·소계 거래처 {code}가 여러 곳에서 중복되었습니다: "
                f"{previous.source.file_name} ↔ {item.source.file_name}. 명세서를 확인해주세요."
            )
        review_vendor_seen[code]=item

    # A verified vendor subtotal is the entire opening balance. Mixing it with
    # separate detail balances for the same vendor would double-count opening.
    overlapping={normalize_code(item.vendor_code) for item in prior_items} & set(review_vendor_seen)
    if overlapping:
        raise ValueError(
            "전월 소계와 개별 명세가 중복될 수 있는 거래처가 있습니다: "
            + ", ".join(sorted(overlapping))
            + ". 중복 명세서/시트를 확인해주세요."
        )

    cross_seen={}
    for item in prior_items:
        key=(item.vendor_code,item.amount,item.description.strip().casefold())
        prev=cross_seen.get(key)
        if prev and prev.source.file_name != item.source.file_name:
            raise ValueError(
                "명세서 간 중복 의심: "
                f"{item.vendor_name or item.vendor_code} / {item.amount:,.0f}원 / "
                f"{prev.source.file_name} ↔ {item.source.file_name}"
            )
        cross_seen.setdefault(key,item)

    journal_items=[]; dz_issues=[]; raw_seen={}
    for path in douzone_paths:
        dz=read_douzone(path,account_codes or None,period)
        source_name=Path(path).name
        for item in dz.items:
            key=(
                item.date,
                normalize_code(item.account_code),
                normalize_code(item.vendor_code),
                normalize_text(item.vendor_name),
                normalize_text(item.description),
                item.debit,
                item.credit,
            )
            prev=raw_seen.get(key)
            if prev and prev != source_name:
                raise ValueError(
                    "더존 Raw 파일 간 동일 전표 의심: "
                    f"{item.date} / {item.vendor_name or item.vendor_code} / "
                    f"차변 {item.debit:,.0f} / 대변 {item.credit:,.0f} / "
                    f"{prev} ↔ {source_name}. 겹치는 Raw 범위를 확인해주세요."
                )
            raw_seen.setdefault(key,source_name)
        journal_items.extend(dz.items); dz_issues.extend(dz.issues)

    issues=prior_issues+dz_issues
    results=reconcile(prior_items,journal_items)
    for item, reason in opening_reviews:
        vendor=normalize_code(item.vendor_code)
        activity=[line for line in journal_items if normalize_code(line.vendor_code)==vendor]
        debits=sum((line.debit for line in activity),start=0)
        credits=sum((line.credit for line in activity),start=0)
        net=item.amount + credits - debits
        reason += (
            f" · 당월 차변 {debits:,.0f}원 / 대변 {credits:,.0f}원"
            f" → 당월말 거래처 순잔액 {net:,.0f}원"
        )
        if issues:
            status=Status.SIGNED_OPENING_REVIEW
            reason += " · 입력 확인 항목이 있어 자동 반영 보류"
        elif net < 0:
            status=Status.SIGNED_OPENING_REVIEW
            reason += " · 음수 잔액(초과 지급/조정) 원인 확인 필요"
        else:
            status=Status.SIGNED_NET_AUTO
            reason += " · 소계 검증 및 거래처 단위 자동계산 완료 (개별 청구 건 배분 미확정)"
        results.append(ReconcileResult(
            status, item, None, reason, "SIGNED_OPENING_NET",closing_balance=net
        ))
    if dz_issues:
        for result in results:
            if result.status in (Status.UNPAID, Status.PARTIAL):
                result.status=Status.RAW_INPUT_INCOMPLETE
                result.reason="더존 대상월/계정 데이터에 자동 제외된 행이 있어 지급 여부를 확정할 수 없음"
                result.rule="RAW_INPUT_INCOMPLETE"
    fresh=new_payables(journal_items, results, allow_auto=not issues)
    counts=Counter(x.status for x in results)
    return ReconciliationRun(period,len(prior_items)+len(opening_reviews),results,fresh,issues,counts)
