"""Independent source totals and ledger comparison, with no Excel dependencies."""
from collections import defaultdict
from decimal import Decimal

from domain.normalization import normalize_code


def build_balance_checks(opening, records, journal, sources, issues):
    vendors={normalize_code(r.vendor_code):r.vendor_name for r in records}
    opening_totals=defaultdict(Decimal)
    output_totals=defaultdict(Decimal)
    activity=defaultdict(list)
    for code,_,_,amount,*_ in opening:
        opening_totals[normalize_code(code)]+=amount
    for row in records:
        output_totals[normalize_code(row.vendor_code)]+=row.amount
    seen=set()
    for line in journal:
        if id(line) not in seen:
            activity[normalize_code(line.vendor_code)].append(line)
            seen.add(id(line))
    checks=[]
    for code,name in vendors.items():
        lines=activity[code]
        credit=sum((x.credit for x in lines),Decimal(0))
        debit=sum((x.debit for x in lines),Decimal(0))
        expected=opening_totals[code]+credit-debit
        output=output_totals[code]
        incomplete=any(not issue.vendor_code or normalize_code(issue.vendor_code)==code for issue in issues)
        calculation=("계산 불일치" if expected!=output else
                     "입력 확인 필요" if incomplete else "계산 일치")
        ledger=None
        reason="원장 잔액 미제공 또는 읽기 불가 · 계산 일치는 원장 일치를 뜻하지 않음"
        accounts=defaultdict(list)
        for line in lines:
            accounts[normalize_code(line.account_code)].append(line)
        if lines and all(x.ledger_balance is not None for x in lines):
            valid=True
            for group in accounts.values():
                # A running balance is meaningful only in a single ordered export.
                if (len({sources.get(id(x),"더존 Raw") for x in group})!=1
                        or [x.date for x in group]!=sorted(x.date for x in group)
                        or any(b.ledger_balance!=a.ledger_balance+b.credit-b.debit
                               for a,b in zip(group,group[1:]))):
                    valid=False
            if valid:
                ledger=sum((group[-1].ledger_balance for group in accounts.values()),Decimal(0))
                reason="제공 RAW 범위의 계정별 마지막 거래 잔액 합계 기준 · 전체 월 원장 제공 여부는 별도 확인"
            else:
                reason="원장 잔액 연결/정렬 또는 파일 분리 확인 필요 · 비교 미확정"
        elif any(x.ledger_balance is not None for x in lines):
            reason="일부 원장 잔액 누락 또는 수식/형식 확인 필요 · 비교 미확정"
        difference=None if ledger is None else output-ledger
        comparison=("비교 미확정" if ledger is None or incomplete else
                    "원장 일치" if difference==0 else "원장 차액 확인")
        checks.append((code,name,opening_totals[code],credit,debit,expected,output,
                       output-expected,calculation,ledger,difference,comparison,reason))
    return checks
