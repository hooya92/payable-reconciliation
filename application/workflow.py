"""UI-independent workflow helpers shared by every front end (Tk and web)."""
from __future__ import annotations

import hashlib
import os
import sys
from collections import Counter
from pathlib import Path

from adapters.excel.reader import detect_douzone_periods, detect_statement_periods, read_douzone, read_prior
from domain.models import Status


def file_digest(path):
    h=hashlib.sha256()
    with open(path,"rb") as f:
        for chunk in iter(lambda:f.read(1024*1024),b""):
            h.update(chunk)
    return h.hexdigest()


def open_saved_folder(path):
    """Show the output directory in Windows Explorer without opening Excel."""
    if sys.platform=="win32":
        os.startfile(str(Path(path).resolve().parent))


def snapshot_file_digests(paths):
    return {str(Path(path).resolve()):file_digest(path) for path in paths}


def changed_snapshot_paths(snapshot):
    changed=[]
    for path,digest in snapshot.items():
        p=Path(path)
        if not p.exists() or file_digest(p)!=digest:
            changed.append(p.name)
    return changed


def find_duplicate_files(prior_paths, douzone_paths, digests=None):
    seen={}; duplicates=[]
    digests=digests or snapshot_file_digests(list(prior_paths)+list(douzone_paths))
    for kind, paths in (("명세서",prior_paths),("더존 원장",douzone_paths)):
        for path in paths:
            digest=digests[str(Path(path).resolve())]
            if digest in seen:
                duplicates.append((seen[digest],(kind,Path(path).name)))
            else:
                seen[digest]=(kind,Path(path).name)
    return duplicates


def reconciliation_breakdown(counts, new_count=0, issue_count=0):
    auto_statuses=(Status.MATCHED,Status.SIGNED_NET_AUTO)
    matched=sum(counts.get(status,0) for status in auto_statuses)
    review=sum(v for k,v in counts.items() if k not in auto_statuses)+new_count
    return {
        "matched":matched,
        "review":review,
        "issues":issue_count,
        "new":new_count,
    }


def completion_detail(period, prior_count, counts, new_count=0, issue_count=0):
    b=reconciliation_breakdown(counts,new_count,issue_count)
    return (
        f"{period.label} · 명세서 {prior_count:,}건 처리 · "
        f"자동 대사 {b['matched']:,}건 · 검토 필요 {b['review']:,}건 · 입력 확인 {b['issues']:,}건"
    )


def cached_periods(cache, kind, path, codes):
    """Months found in a file. Reading a big workbook is slow, so remember the answer until
    the file (size / modification time) or, for the ledger, the account codes change."""
    stat=os.stat(path)
    key=(kind,str(path),stat.st_size,stat.st_mtime_ns,tuple(sorted(codes)) if kind=="douzone" else ())
    if key not in cache:
        cache[key]=(detect_statement_periods(path) if kind=="prior"
                    else detect_douzone_periods(path,codes))
    return cache[key]


def suggest_reconciliation_period(current_period, statement_period_groups):
    """Infer the target accounting month directly from the latest statement month."""
    groups=[set(group) for group in statement_period_groups if group]
    if not groups:
        return None
    common=set.intersection(*groups)
    if not common:
        return None
    return max(common).next()


def infer_period_from_inputs(statement_period_groups, raw_period_groups, preferred=None):
    """Infer the target month without falling back to an older coincidental overlap."""
    statement_groups=[set(group) for group in statement_period_groups if group]
    raw_groups=[set(group) for group in raw_period_groups if group]

    statement_common=set.intersection(*statement_groups) if statement_groups else set()
    raw_periods=set().union(*raw_groups) if raw_groups else set()

    if statement_common:
        target=max(statement_common).next()
        if preferred is not None and preferred.previous() in statement_common and (
            not raw_periods or preferred in raw_periods
        ):
            target=preferred
        if raw_periods:
            return (target,"statement+raw") if target in raw_periods else (None,"conflict")
        return target,"statement"
    if raw_periods:
        # A manually selected month must not be replaced by the latest month
        # in an annual Raw when the statement tab has no readable month.
        if preferred is not None and preferred in raw_periods:
            return preferred,"raw"
        return max(raw_periods),"raw"
    return None,"unknown"


def period_alignment_evidence(prior_paths, douzone_paths, account_codes, period):
    """Return (exact code+amount matches, shared vendor codes) for a candidate month."""
    prior_pairs=Counter()
    raw_pairs=Counter()
    prior_codes=set()
    raw_codes=set()
    try:
        for path in prior_paths:
            rr=read_prior(path,Path(path).stem,period.previous())
            for item in rr.items:
                prior_pairs[(item.vendor_code,item.amount)]+=1
                prior_codes.add(item.vendor_code)
        for path in douzone_paths:
            rr=read_douzone(path,account_codes or None,period)
            for item in rr.items:
                if item.debit>0:
                    raw_pairs[(item.vendor_code,item.debit)]+=1
                    raw_codes.add(item.vendor_code)
    except Exception:
        return 0,0

    exact=sum(min(count,raw_pairs.get(key,0)) for key,count in prior_pairs.items())
    shared_codes=len(prior_codes & raw_codes)
    return exact,shared_codes
