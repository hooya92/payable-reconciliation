from __future__ import annotations
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from openpyxl import load_workbook

from domain.models import DataQuality, JournalLine, PayableItem, SourceRef
from domain.normalization import normalize_code, parse_amount
from domain.period import AccountingPeriod, parse_date


@dataclass(frozen=True)
class InputIssue:
    source: str
    sheet: str
    row: int
    field: str
    raw_value: object
    reason: str


@dataclass
class ReadResult:
    items: list = field(default_factory=list)
    issues: list[InputIssue] = field(default_factory=list)
    detected_headers: list[str] = field(default_factory=list)


def _text(v): return "" if v is None else str(v).strip()
def _header(v): return "".join(_text(v).lower().split())
def _date(v):
    if isinstance(v, (datetime, date)): return v.strftime("%Y-%m-%d")
    return _text(v)


def _header_row(ws, required_groups, max_rows=30):
    for r in range(1, min(ws.max_row, max_rows) + 1):
        vals=[_header(ws.cell(r,c).value) for c in range(1,ws.max_column+1)]
        if all(any(name in vals for name in group) for group in required_groups):
            return r, vals
    return None, []


def _first(vals, *names):
    for name in names:
        try: return vals.index(_header(name))+1
        except ValueError: pass
    return None


def _douzone_columns(vals):
    # 더존 전표출력은 '코드'가 2개일 수 있다.
    # 계정과목명 바로 왼쪽 코드=계정코드, 거래처명 바로 왼쪽 코드=거래처코드.
    account_name=_first(vals,"계정과목명","계정과목","계정명")
    vendor_name=_first(vals,"거래처명","거래처 명")
    account_code=_first(vals,"계정코드","계정 코드")
    vendor_code=_first(vals,"거래처코드","거래처 코드")
    if account_code is None and account_name and account_name>1 and vals[account_name-2]=="코드":
        account_code=account_name-1
    if vendor_code is None and vendor_name and vendor_name>1 and vals[vendor_name-2]=="코드":
        vendor_code=vendor_name-1
    return {
        "date":_first(vals,"기표일자","일자","날짜"),
        "account_code":account_code, "account_name":account_name,
        "vendor_code":vendor_code, "vendor_name":vendor_name,
        "description":_first(vals,"적요","적요명"),
        "debit":_first(vals,"차변","차변금액","차변 금액"),
        "credit":_first(vals,"대변","대변금액","대변 금액"),
    }


def read_douzone(path: str|Path, account_codes:set[str]|None=None, period: AccountingPeriod|None=None) -> ReadResult:
    wb=load_workbook(path,read_only=True,data_only=True)
    out=ReadResult(); wanted={normalize_code(x) for x in (account_codes or set()) if x}
    for ws in wb.worksheets:
        hr, vals=_header_row(ws,[("거래처명","거래처 명"),("적요","적요명"),("차변","차변금액","차변 금액"),("대변","대변금액","대변 금액")])
        if not hr: continue
        cols=_douzone_columns(vals)
        needed=("vendor_code","vendor_name","description","debit","credit")
        missing=[x for x in needed if not cols[x]]
        if missing:
            raise ValueError("더존 파일 필수 헤더를 안전하게 식별하지 못했습니다: "+", ".join(missing))
        out.detected_headers=[_text(ws.cell(hr,c).value) for c in range(1,ws.max_column+1)]
        for r in range(hr+1,ws.max_row+1):
            raw_date=ws.cell(r,cols["date"]).value if cols["date"] else None
            if period:
                parsed_date=parse_date(raw_date)
                if parsed_date is None:
                    out.issues.append(InputIssue("더존",ws.title,r,"기표일자",raw_date,"날짜를 해석할 수 없어 대상월 필터에서 제외")); continue
                if not period.contains(parsed_date): continue
            ac=normalize_code(ws.cell(r,cols["account_code"]).value) if cols["account_code"] else ""
            an=_text(ws.cell(r,cols["account_name"]).value) if cols["account_name"] else ""
            if wanted and ac not in wanted: continue
            if not wanted and an and "미지급" not in an: continue
            d=parse_amount(ws.cell(r,cols["debit"]).value); cr=parse_amount(ws.cell(r,cols["credit"]).value)
            bad=[("차변",d),("대변",cr)]
            suspicious=False
            for field,x in bad:
                if x.quality==DataQuality.SUSPICIOUS:
                    out.issues.append(InputIssue("더존",ws.title,r,field,x.raw,x.reason)); suspicious=True
            if suspicious: continue
            debit=d.value or 0; credit=cr.value or 0
            if debit==0 and credit==0: continue
            code=normalize_code(ws.cell(r,cols["vendor_code"]).value)
            if not code:
                out.issues.append(InputIssue("더존",ws.title,r,"거래처코드","", "거래처코드 없음")); continue
            out.items.append(JournalLine(code,_text(ws.cell(r,cols["vendor_name"]).value),ac,an,
                _text(ws.cell(r,cols["description"]).value),debit,credit,
                _date(raw_date) if cols["date"] else "",r))
    wb.close()
    if not out.items and not out.issues: raise ValueError("더존 파일에서 대상 미지급금 전표를 찾지 못했습니다.")
    return out


def read_prior(path: str|Path, owner: str = "") -> ReadResult:
    wb=load_workbook(path,read_only=True,data_only=True); out=ReadResult()
    for ws in wb.worksheets:
        hr,vals=_header_row(ws,[("거래처코드","거래처 코드","코드"),("거래처명","거래처","업체명"),("적요","내역","내용"),("금액","미지급금","잔액")])
        if not hr: continue
        cols={
            "vendor_code":_first(vals,"거래처코드","거래처 코드","코드"),
            "vendor_name":_first(vals,"거래처명","거래처","업체명"),
            "date":_first(vals,"날짜","일자","기표일자"),
            "description":_first(vals,"적요","내역","내용"),
            "amount":_first(vals,"금액","미지급금","잔액")}
        out.detected_headers=[_text(ws.cell(hr,c).value) for c in range(1,ws.max_column+1)]
        for r in range(hr+1,ws.max_row+1):
            code=normalize_code(ws.cell(r,cols["vendor_code"]).value)
            desc=_text(ws.cell(r,cols["description"]).value)
            a=parse_amount(ws.cell(r,cols["amount"]).value)
            if not code and not desc and (a.value or 0)==0: continue  # subtotal/blank
            if a.quality==DataQuality.SUSPICIOUS:
                out.issues.append(InputIssue("전월명세",ws.title,r,"금액",a.raw,a.reason)); continue
            if not code or not desc or (a.value or 0)<=0:
                continue
            out.items.append(PayableItem(code,_text(ws.cell(r,cols["vendor_name"]).value),desc,a.value,
                _date(ws.cell(r,cols["date"]).value) if cols["date"] else "",r, SourceRef(Path(path).name, ws.title, r, owner)))
    wb.close()
    if not out.items and not out.issues: raise ValueError("전월 명세서에서 대사할 항목을 찾지 못했습니다.")
    return out
