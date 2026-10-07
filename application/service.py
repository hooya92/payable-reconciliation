from collections import Counter
from dataclasses import dataclass
from pathlib import Path

from adapters.excel.reader import read_douzone, read_prior
from domain.models import Status
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
    prior_items=[]; prior_issues=[]
    for path in prior_paths:
        rr=read_prior(path,Path(path).stem,period)
        prior_items.extend(rr.items); prior_issues.extend(rr.issues)

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
    if dz_issues:
        for result in results:
            if result.status==Status.UNPAID:
                result.status=Status.RAW_INPUT_INCOMPLETE
                result.reason="더존 대상월/계정 데이터에 자동 제외된 행이 있어 지급 여부를 확정할 수 없음"
                result.rule="RAW_INPUT_INCOMPLETE"
    fresh=new_payables(journal_items)
    counts=Counter(x.status for x in results)
    return ReconciliationRun(period,len(prior_items),results,fresh,issues,counts)
