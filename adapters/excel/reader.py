from __future__ import annotations
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
import re
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
    recognized_sheets: list[str] = field(default_factory=list)


def _text(v): return "" if v is None else str(v).strip()
def _is_formula(cell): return getattr(cell,"data_type",None)=="f" or (isinstance(cell.value,str) and cell.value.startswith("="))
def _code_from_cell(cell):
    """Preserve identifier leading zeroes when Excel stores a numeric code with a zero-only display format."""
    value=cell.value
    if value is None: return ""
    if isinstance(value,(int,float)) and not isinstance(value,bool):
        fmt=(cell.number_format or "").strip()
        if re.fullmatch(r"0+",fmt):
            if isinstance(value,float) and not value.is_integer():
                return normalize_code(value)
            return str(int(value)).zfill(len(fmt))
    return normalize_code(value)
def _header(v): return "".join(_text(v).lower().split())
def _date(v):
    if isinstance(v, (datetime, date)): return v.strftime("%Y-%m-%d")
    return _text(v)
def _is_summary_label(v):
    return _header(v) in {"소계","합계","총계","전체합계"}


def _header_row(ws, required_groups, max_rows=30):
    for r in range(1, min(ws.max_row, max_rows) + 1):
        vals=[_header(ws.cell(r,c).value) for c in range(1,ws.max_column+1)]
        if all(any(name in vals for name in group) for group in required_groups):
            return r, vals
    return None, []


def _partial_header_row(ws, required_groups, min_groups, max_rows=30):
    """Detect a likely data table whose headers changed enough that we should not silently ignore it."""
    best=(None,0)
    for r in range(1, min(ws.max_row, max_rows) + 1):
        vals=[_header(ws.cell(r,c).value) for c in range(1,ws.max_column+1)]
        matched=sum(any(_header(name) in vals for name in group) for group in required_groups)
        if matched>best[1]:
            best=(r,matched)
    return best if best[1]>=min_groups else (None,0)


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
    wanted={normalize_code(x) for x in (account_codes or set()) if x}
    if not wanted:
        raise ValueError("미지급금 계정코드를 1개 이상 지정해야 합니다.")
    wb=load_workbook(path,read_only=True,data_only=False)
    out=ReadResult()
    groups=[("거래처명","거래처 명"),("적요","적요명"),("차변","차변금액","차변 금액"),("대변","대변금액","대변 금액")]
    try:
        for ws in wb.worksheets:
            hr, vals=_header_row(ws,groups)
            if not hr:
                partial_row,matched=_partial_header_row(ws,groups,3)
                if partial_row:
                    raise ValueError(
                        f"더존 파일의 '{ws.title}' 시트가 전표 표처럼 보이지만 필수 헤더를 완전히 식별하지 못했습니다 "
                        f"(필수 항목 {matched}/4개 인식). 시트 구조를 확인해주세요."
                    )
                continue
            out.recognized_sheets.append(ws.title)
            cols=_douzone_columns(vals)
            needed=["vendor_code","vendor_name","description","debit","credit","account_code"]
            if period: needed.append("date")
            missing=[x for x in needed if not cols[x]]
            if missing:
                raise ValueError("더존 파일 필수 헤더를 안전하게 식별하지 못했습니다: "+", ".join(missing))
            out.detected_headers=[_text(ws.cell(hr,c).value) for c in range(1,ws.max_column+1)]
            for r, row in enumerate(ws.iter_rows(min_row=hr+1, values_only=False), start=hr+1):
                account_cell=row[cols["account_code"]-1]
                if _is_formula(account_cell):
                    out.issues.append(InputIssue("더존",ws.title,r,"계정코드",account_cell.value,"수식 셀은 대상 계정 여부를 확정할 수 없어 자동 처리하지 않음")); continue
                ac=_code_from_cell(account_cell)
                if ac not in wanted:
                    continue
                an=_text(row[cols["account_name"]-1].value) if cols["account_name"] else ""

                date_cell=row[cols["date"]-1] if cols["date"] else None
                raw_date=date_cell.value if date_cell else None
                if date_cell is not None and _is_formula(date_cell):
                    out.issues.append(InputIssue("더존",ws.title,r,"기표일자",raw_date,"수식 날짜는 대상월을 확정할 수 없어 자동 처리하지 않음")); continue
                if period:
                    parsed_date=parse_date(raw_date)
                    if parsed_date is None:
                        out.issues.append(InputIssue("더존",ws.title,r,"기표일자",raw_date,"날짜를 해석할 수 없어 대상월 필터에서 제외")); continue
                    if not period.contains(parsed_date): continue

                formula_field=None
                for field,col in (("거래처코드",cols["vendor_code"]),("거래처명",cols["vendor_name"]),
                                  ("적요",cols["description"]),("차변",cols["debit"]),("대변",cols["credit"])):
                    cell=row[col-1]
                    if _is_formula(cell):
                        formula_field=(field,cell.value); break
                if formula_field:
                    out.issues.append(InputIssue("더존",ws.title,r,formula_field[0],formula_field[1],
                        "수식 셀은 계산값의 최신성을 보장할 수 없어 자동 처리하지 않음")); continue

                d=parse_amount(row[cols["debit"]-1].value); cr=parse_amount(row[cols["credit"]-1].value)
                suspicious=False
                for field,x in (("차변",d),("대변",cr)):
                    if x.quality==DataQuality.SUSPICIOUS:
                        out.issues.append(InputIssue("더존",ws.title,r,field,x.raw,x.reason)); suspicious=True
                if suspicious: continue

                debit=d.value or 0; credit=cr.value or 0
                if debit==0 and credit==0: continue
                if debit < 0 or credit < 0:
                    out.issues.append(InputIssue("더존",ws.title,r,"차변/대변",
                        f"{debit}/{credit}","음수 전표는 수정·역분개 가능성이 있어 자동 처리하지 않음")); continue
                if debit > 0 and credit > 0:
                    out.issues.append(InputIssue("더존",ws.title,r,"차변/대변",
                        f"{debit}/{credit}","한 행에 차변과 대변이 동시에 있어 자동 처리하지 않음")); continue

                code=_code_from_cell(row[cols["vendor_code"]-1])
                vendor_name=_text(row[cols["vendor_name"]-1].value)
                desc=_text(row[cols["description"]-1].value)
                row_invalid=False
                if not code:
                    out.issues.append(InputIssue("더존",ws.title,r,"거래처코드","", "거래처코드 없음")); row_invalid=True
                if not vendor_name:
                    out.issues.append(InputIssue("더존",ws.title,r,"거래처명","", "거래처명 없음")); row_invalid=True
                if not desc:
                    out.issues.append(InputIssue("더존",ws.title,r,"적요","", "적요 없음")); row_invalid=True
                if row_invalid: continue

                out.items.append(JournalLine(code,vendor_name,ac,an,desc,debit,credit,
                    _date(raw_date) if cols["date"] else "",r))
    finally:
        wb.close()
    if not out.items and not out.issues: raise ValueError("더존 파일에서 대상 미지급금 전표를 찾지 못했습니다.")
    return out


def read_prior(path: str|Path, owner: str = "") -> ReadResult:
    wb=load_workbook(path,read_only=True,data_only=False); out=ReadResult()
    groups=[("거래처코드","거래처 코드","코드"),("거래처명","거래처","업체명"),("적요","내역","내용"),("금액","미지급금","잔액")]
    try:
        for ws in wb.worksheets:
            hr,vals=_header_row(ws,groups)
            if not hr:
                partial_row,matched=_partial_header_row(ws,groups,3)
                if partial_row:
                    raise ValueError(
                        f"명세서의 '{ws.title}' 시트가 데이터 표처럼 보이지만 필수 헤더를 완전히 식별하지 못했습니다 "
                        f"(필수 항목 {matched}/4개 인식). 시트 구조를 확인해주세요."
                    )
                continue
            out.recognized_sheets.append(ws.title)
            cols={
                "vendor_code":_first(vals,"거래처코드","거래처 코드","코드"),
                "vendor_name":_first(vals,"거래처명","거래처","업체명"),
                "date":_first(vals,"날짜","일자","기표일자"),
                "description":_first(vals,"적요","내역","내용"),
                "amount":_first(vals,"금액","미지급금","잔액")}
            required=("vendor_code","vendor_name","description","amount")
            missing=[x for x in required if not cols[x]]
            if missing:
                raise ValueError("전월 명세서 필수 헤더를 안전하게 식별하지 못했습니다: "+", ".join(missing))
            out.detected_headers=[_text(ws.cell(hr,c).value) for c in range(1,ws.max_column+1)]
            for r, row in enumerate(ws.iter_rows(min_row=hr+1, values_only=False), start=hr+1):
                code_cell=row[cols["vendor_code"]-1]
                name_cell=row[cols["vendor_name"]-1]
                desc_cell=row[cols["description"]-1]
                amount_cell=row[cols["amount"]-1]
                code=_code_from_cell(code_cell)
                vendor_name=_text(name_cell.value)
                desc=_text(desc_cell.value)

                if not code and not desc and _is_summary_label(vendor_name):
                    continue
                if any(_is_formula(cell) for cell in (code_cell,name_cell,desc_cell,amount_cell)):
                    out.issues.append(InputIssue("전월명세",ws.title,r,"수식",amount_cell.value,
                        "수식 셀은 계산값의 최신성을 보장할 수 없어 자동 대사하지 않음")); continue

                a=parse_amount(amount_cell.value)
                if not code and not vendor_name and not desc and (a.value or 0)==0:
                    continue
                if a.quality==DataQuality.SUSPICIOUS:
                    out.issues.append(InputIssue("전월명세",ws.title,r,"금액",a.raw,a.reason)); continue

                row_invalid=False
                if not code:
                    out.issues.append(InputIssue("전월명세",ws.title,r,"거래처코드","", "거래처코드 없음")); row_invalid=True
                if not vendor_name:
                    out.issues.append(InputIssue("전월명세",ws.title,r,"거래처명","", "거래처명 없음")); row_invalid=True
                if not desc:
                    out.issues.append(InputIssue("전월명세",ws.title,r,"적요","", "적요 없음")); row_invalid=True
                if (a.value or 0) <= 0:
                    out.issues.append(InputIssue("전월명세",ws.title,r,"금액",a.raw,"금액이 0 이하라 자동 대사하지 않음")); row_invalid=True
                if row_invalid: continue

                out.items.append(PayableItem(code,vendor_name,desc,a.value,
                    _date(row[cols["date"]-1].value) if cols["date"] else "",r, SourceRef(Path(path).name, ws.title, r, owner)))
    finally:
        wb.close()
    if not out.items and not out.issues: raise ValueError("전월 명세서에서 대사할 항목을 찾지 못했습니다.")
    return out
