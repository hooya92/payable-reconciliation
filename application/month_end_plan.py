"""Pure monthly statement plan: source rows, signed Douzone entries, and review history.

No Excel/openpyxl dependency. The writer only applies rows and native styles.
"""
from dataclasses import dataclass
from typing import NamedTuple
from decimal import Decimal

from domain.models import Status
from domain.normalization import normalize_code, normalize_text


class StatementRow(NamedTuple):
    """Stable interface between reconciliation planning and Excel rendering.

    NamedTuple preserves backwards-compatible tuple access in older tests
    while making each field's meaning explicit at module boundaries.
    """
    vendor_code: str
    vendor_name: str
    description: str
    amount: Decimal
    date: object
    status: str
    source_row: int
    extras: dict


@dataclass
class MonthEndPlan:
    records: list[StatementRow]
    audits: list
    offsets: list
    reviews: list
    raw_count: int


def _status_for_opening(result):
    if result is None:
        return "확인 필요"
    if result.status == Status.MATCHED:
        return ("확인 필요" if result.rule=="CODE_AMOUNT_UNIQUE_WITH_NOTE"
                else "대사 일치")
    if result.status == Status.UNPAID:
        return "전월 이월"
    if result.status == Status.INPUT_TYPO_SUSPECT:
        return "오타 의심"
    if result.status == Status.SIGNED_NET_AUTO:
        return "순잔액 대사"
    return "확인 필요"


def _amount(value):
    value=Decimal(value)
    return int(value) if value==value.to_integral_value() else float(value)


def build_month_end_plan(opening, results, new_items, issues, source_name, source_sheet,
                         standalone_debits=(), standalone_debit_sources=None,
                         journal_items=(), journal_sources=None):
    """Keep original detail history, then append each selected current RAW entry once.

    A debit is negative in the one-column statement; a credit is positive.
    This is a draft, not a claim that uncertain allocations were approved.
    """
    original_result_rows={
        result.prior.source.row:result for result in results
        if result.prior.source.file_name==source_name
        and result.prior.source.sheet==source_sheet
    }
    net_vendors={
        normalize_code(result.prior.vendor_code):result for result in results
        if result.status in (Status.SIGNED_NET_AUTO,Status.SIGNED_OPENING_REVIEW)
    }
    records=[]
    audits=[]
    offsets=[]
    reviews=[]
    vendor_order={}
    opening_vendor_by_code={}
    for code,name,desc,amount,when,source_row,extras in opening:
        vendor=(code,name)
        normalized=normalize_code(code)
        previous=opening_vendor_by_code.setdefault(normalized,vendor)
        if normalize_text(previous[1])!=normalize_text(name):
            raise ValueError(
                f"전월 명세서에 동일 거래처코드 {code}의 거래처명이 서로 다릅니다. "
                "행을 자동 병합하지 않고 확인을 요청합니다."
            )
        vendor_order.setdefault(vendor,len(vendor_order))
        result=original_result_rows.get(source_row)
        if result is None:
            result=net_vendors.get(normalize_code(code))
        status=_status_for_opening(result)
        records.append(StatementRow(code,name,desc,amount,when,status,source_row,extras))
        audits.append((status,code,name,desc,_amount(amount),
                       "전월 명세서 원본 행 유지",source_name,source_row))
        if status in ("확인 필요","오타 의심"):
            reason=(result.reason if result else "전월 원본 상세 행과 대사결과 연결 확인 필요")
            reviews.append((status,code,name,desc,_amount(amount),
                            reason,source_name,source_row))

    # Each accepted period/account-code Douzone RAW line is appended once.
    # The status is informational, never a declaration of payment approval.
    source_files=journal_sources if journal_sources is not None else (standalone_debit_sources or {})
    matched_debits={}
    review_debits={}
    for result in results:
        if result.journal is None:
            continue
        if result.status==Status.MATCHED and result.rule!="CODE_AMOUNT_UNIQUE_WITH_NOTE":
            matched_debits[id(result.journal)]=result
            p=result.prior
            j=result.journal
            offsets.append(("대사 일치 후보",p.vendor_code,p.vendor_name,p.description,
                            _amount(p.amount),p.date,p.source.file_name,p.source.row,
                            j.description,j.date,_amount(j.debit),None,j.row_number))
        else:
            review_debits[id(result.journal)]=result
    fresh_by_id={id(item.journal):item for item in new_items}
    standalone_ids={id(line) for line in standalone_debits}
    # One vendor-code group must stay one group, even if separate RAW exports
    # use inconsistent names. The first name becomes display text; disagreement
    # is explicitly listed for review, never silently approved.
    current_vendor_by_code={}
    seen=set()
    sources=list(journal_items)
    if not sources:
        sources=([r.journal for r in results if r.journal is not None]
                 +[item.journal for item in new_items]+list(standalone_debits))
    for line in sources:
        if id(line) in seen:
            continue
        seen.add(id(line))
        if line.credit>0 and line.debit>0:
            raise ValueError("더존 RAW 전표에 차변과 대변이 동시에 있습니다.")
        if line.credit==0 and line.debit==0:
            continue
        if line.credit>0:
            amount=line.credit
            new=fresh_by_id.get(id(line))
            status="당월 발생" if new and new.auto_carry else "확인 필요"
            reason=(new.reason if new else "신규 발생의 대사 조건을 확인해주세요.")
        else:
            amount=-line.debit
            if id(line) in matched_debits:
                status="대사 일치"
                reason=matched_debits[id(line)].reason
            elif id(line) in review_debits:
                result=review_debits[id(line)]
                status=("오타 의심" if result.status==Status.INPUT_TYPO_SUSPECT
                        else "확인 필요")
                reason=result.reason
            elif id(line) in standalone_ids:
                status="원장 단독"
                reason="전월 명세서 대응 건 없음 · RAW 차변 그대로 반영"
            else:
                status="당월 지급"
                reason="더존 RAW 차변 · 개별 전월 발생건과 연결 미확정"
        # Keep the original vendor grouping even if the RAW spells the name
        # differently. Do not silently approve the name discrepancy.
        source_vendor=(line.vendor_code,line.vendor_name)
        vendor_code=normalize_code(line.vendor_code)
        vendor=current_vendor_by_code.setdefault(
            vendor_code,opening_vendor_by_code.get(vendor_code,source_vendor)
        )
        if normalize_text(vendor[1]) != normalize_text(line.vendor_name):
            status="확인 필요"
            reason += f" · 원장 거래처명 '{line.vendor_name}' ↔ 명세서 '{vendor[1]}' 확인"
        vendor_order.setdefault(vendor,len(vendor_order))
        records.append(StatementRow(vendor[0],vendor[1],line.description,
                                    amount,line.date,status,line.row_number,{}))
        audits.append((status,line.vendor_code,line.vendor_name,line.description,
                       _amount(amount),reason,
                       source_files.get(id(line),"더존 Raw"),line.row_number))
        if status in ("확인 필요","오타 의심","원장 단독"):
            reviews.append((status,line.vendor_code,line.vendor_name,line.description,
                            _amount(amount),reason,
                            source_files.get(id(line),"더존 Raw"),line.row_number))

    for issue in issues:
        reviews.append(("입력 데이터 확인","","",str(issue.raw_value),"",
                        issue.reason,issue.source,issue.row))
    records.sort(key=lambda record:vendor_order[(record.vendor_code,record.vendor_name)])
    return MonthEndPlan(records,audits,offsets,reviews,len(seen))
