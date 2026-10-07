from collections import Counter
from dataclasses import dataclass
from pathlib import Path

from adapters.excel.reader import read_douzone, read_prior
from domain.models import Status
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
    prior_items=[]; prior_issues=[]
    for path in prior_paths:
        rr=read_prior(path,Path(path).stem)
        prior_items.extend(rr.items); prior_issues.extend(rr.issues)

    cross_seen={}
    for item in prior_items:
        key=(item.vendor_code,int(item.amount),item.description.strip().casefold())
        prev=cross_seen.get(key)
        if prev and prev.source.file_name != item.source.file_name:
            raise ValueError(
                "명세서 간 중복 의심: "
                f"{item.vendor_name or item.vendor_code} / {int(item.amount):,}원 / "
                f"{prev.source.file_name} ↔ {item.source.file_name}"
            )
        cross_seen.setdefault(key,item)

    journal_items=[]; dz_issues=[]
    for path in douzone_paths:
        dz=read_douzone(path,account_codes or None,period)
        journal_items.extend(dz.items); dz_issues.extend(dz.issues)

    issues=prior_issues+dz_issues
    results=reconcile(prior_items,journal_items)
    fresh=new_payables(journal_items)
    counts=Counter(x.status for x in results)
    return ReconciliationRun(period,len(prior_items),results,fresh,issues,counts)
