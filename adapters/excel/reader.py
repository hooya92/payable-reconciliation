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


def _sheet_matches_period(title, period):
    """Recognize common monthly sheet names such as 26.07, 2026-07, 2026년 7월."""
    if period is None:
        return False
    text=_text(title).strip()
    year=str(period.year); yy=year[-2:]; month=str(period.month); mm=f"{period.month:02d}"
    patterns=(
        rf"(?<!\d){re.escape(year)}[.\-_/]{re.escape(mm)}(?!\d)",
        rf"(?<!\d){re.escape(yy)}[.\-_/]{re.escape(mm)}(?!\d)",
        rf"(?<!\d){re.escape(year)}\s*년\s*0?{re.escape(month)}\s*월",
        rf"(?<!\d){re.escape(yy)}\s*년\s*0?{re.escape(month)}\s*월",
    )
    return any(re.search(pattern,text,re.IGNORECASE) for pattern in patterns)


def _select_prior_sheets(wb, groups, period):
    """Pick only the intended monthly statement sheet when a workbook contains many month tabs."""
    full=[]
    for ws in wb.worksheets:
        hr,vals=_header_row(ws,groups)
        if hr:
            full.append((ws,hr,vals))

    if period is not None:
        title_matches=[entry for entry in full if _sheet_matches_period(entry[0].title,period)]
        if len(title_matches)==1:
            return title_matches
        if len(title_matches)>1:
            names=", ".join(x[0].title for x in title_matches)
            raise ValueError(f"{period.label} 명세서 시트가 여러 개라 자동 선택할 수 없습니다: {names}")
        if len(full)>1:
            names=", ".join(x[0].title for x in full)
            raise ValueError(
                f"명세서 파일에 월별 데이터 시트가 여러 개 있지만 {period.label} 시트를 찾지 못했습니다. "
                f"확인된 시트: {names}"
            )

    if len(full)==1:
        return full
    if len(full)>1:
        names=", ".join(x[0].title for x in full)
        raise ValueError(f"명세서 데이터 시트가 여러 개라 자동 선택할 수 없습니다: {names}")

    # No full header was found. A data-like partial header must not be silently ignored.
    for ws in wb.worksheets:
        partial_row,matched=_partial_header_row(ws,groups,3)
        if partial_row:
            raise ValueError(
                f"명세서의 '{ws.title}' 시트가 데이터 표처럼 보이지만 필수 헤더를 완전히 식별하지 못했습니다 "
                f"(필수 항목 {matched}/4개 인식). 시트 구조를 확인해주세요."
            )
    return []


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


def read_prior(path: str|Path, owner: str = "", period: AccountingPeriod|None=None) -> ReadResult:
    wb=load_workbook(path,read_only=True,data_only=False); out=ReadResult()
    groups=[("거래처코드","거래처 코드","코드"),("거래처명","거래처","업체명"),("적요","내역","내용"),("금액","미지급금","잔액")]
    try:
        selected=_select_prior_sheets(wb,groups,period)
        for ws,hr,vals in selected:
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

            # Some real statement sheets place vendor code/name only on the yellow subtotal row
            # after the detail rows. Buffer only contiguous detail rows and attach the vendor
            # identity when the following subtotal exactly equals their amount sum.
            pending=[]
            def flush_pending_as_issues(reason):
                nonlocal pending
                for rec in pending:
                    out.issues.append(InputIssue("전월명세",ws.title,rec["row"],"거래처",rec["description"],reason))
                pending=[]

            for r,row in enumerate(ws.iter_rows(min_row=hr+1,values_only=False),start=hr+1):
                code_cell=row[cols["vendor_code"]-1]
                name_cell=row[cols["vendor_name"]-1]
                desc_cell=row[cols["description"]-1]
                amount_cell=row[cols["amount"]-1]
                date_cell=row[cols["date"]-1] if cols["date"] else None

                code=_code_from_cell(code_cell)
                vendor_name=_text(name_cell.value)
                desc=_text(desc_cell.value)
                raw_date=date_cell.value if date_cell else None
                summary_label=next(
                    (v for v in (raw_date,desc,vendor_name) if _is_summary_label(v)),
                    None,
                )

                relevant_cells=[code_cell,name_cell,desc_cell,amount_cell]
                if date_cell is not None:
                    relevant_cells.append(date_cell)
                if any(_is_formula(cell) for cell in relevant_cells):
                    if pending:
                        flush_pending_as_issues("그룹 중간에 수식 셀이 있어 거래처를 안전하게 확정할 수 없음")
                    out.issues.append(InputIssue("전월명세",ws.title,r,"수식",amount_cell.value,
                        "수식 셀은 계산값의 최신성을 보장할 수 없어 자동 대사하지 않음"))
                    continue

                a=parse_amount(amount_cell.value)
                if not code and not vendor_name and not desc and (a.value or 0)==0:
                    if pending:
                        flush_pending_as_issues("거래처 소계 전에 빈 행이 있어 그룹 범위를 확정할 수 없음")
                    continue
                if a.quality==DataQuality.SUSPICIOUS:
                    if pending:
                        flush_pending_as_issues("그룹 중간에 확인이 필요한 금액 형식이 있어 거래처를 확정할 수 없음")
                    out.issues.append(InputIssue("전월명세",ws.title,r,"금액",a.raw,a.reason))
                    continue

                # Group subtotal row: vendor code/name + '소계' and exact sum of preceding details.
                if code and vendor_name and summary_label:
                    subtotal=a.value or 0
                    if subtotal <= 0:
                        flush_pending_as_issues("거래처 소계 금액이 0 이하라 그룹을 확정할 수 없음")
                        out.issues.append(InputIssue("전월명세",ws.title,r,"금액",a.raw,"소계 금액이 0 이하"))
                        continue
                    if not pending:
                        # A subtotal without detail rows is presentation-only; do not turn it into a payable.
                        continue
                    pending_total=sum(x["amount"] for x in pending)
                    if pending_total != subtotal:
                        reason=f"상세 합계 {pending_total:,.0f}원과 소계 {subtotal:,.0f}원이 달라 자동 대사하지 않음"
                        flush_pending_as_issues(reason)
                        out.issues.append(InputIssue("전월명세",ws.title,r,"소계",a.raw,reason))
                        continue
                    for rec in pending:
                        out.items.append(PayableItem(
                            code,vendor_name,rec["description"],rec["amount"],
                            rec["date"],rec["row"],SourceRef(Path(path).name,ws.title,rec["row"],owner)
                        ))
                    pending=[]
                    continue

                # Flat format: every detail row already carries vendor identity.
                if code or vendor_name:
                    if pending:
                        flush_pending_as_issues("거래처 소계가 나오기 전에 다른 거래처 행이 시작되어 그룹을 확정할 수 없음")
                    row_invalid=False
                    if not code:
                        out.issues.append(InputIssue("전월명세",ws.title,r,"거래처코드","", "거래처코드 없음")); row_invalid=True
                    if not vendor_name:
                        out.issues.append(InputIssue("전월명세",ws.title,r,"거래처명","", "거래처명 없음")); row_invalid=True
                    if not desc:
                        out.issues.append(InputIssue("전월명세",ws.title,r,"적요","", "적요 없음")); row_invalid=True
                    if (a.value or 0) <= 0:
                        out.issues.append(InputIssue("전월명세",ws.title,r,"금액",a.raw,"금액이 0 이하라 자동 대사하지 않음")); row_invalid=True
                    if row_invalid:
                        continue
                    out.items.append(PayableItem(
                        code,vendor_name,desc,a.value,_date(raw_date) if date_cell else "",r,
                        SourceRef(Path(path).name,ws.title,r,owner)
                    ))
                    continue

                # Grouped format detail row: description+amount are present, vendor is supplied by
                # the following subtotal row. Keep it buffered; never guess the vendor.
                if desc and (a.value or 0)>0:
                    pending.append({
                        "row":r,
                        "date":_date(raw_date) if date_cell else "",
                        "description":desc,
                        "amount":a.value,
                    })
                    continue

                if not desc and (a.value or 0)>0:
                    if pending:
                        flush_pending_as_issues("적요 없는 행이 그룹 중간에 있어 거래처를 확정할 수 없음")
                    out.issues.append(InputIssue("전월명세",ws.title,r,"적요","", "적요 없음"))
                    continue

            if pending:
                flush_pending_as_issues("파일 끝까지 거래처 소계가 없어 상세 행의 거래처를 확정할 수 없음")
    finally:
        wb.close()
    if not out.items and not out.issues:
        target=f" {period.label}" if period else ""
        raise ValueError(f"전월 명세서에서{target} 대사할 항목을 찾지 못했습니다.")
    return out

