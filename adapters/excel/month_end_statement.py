"""Create an editable month-end statement using the previous statement as a visual template.

The input workbook is never saved. The main sheet keeps the template columns;
review explanations and change history live on removable companion sheets.
"""
from __future__ import annotations

from copy import copy
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
import calendar
import re

from openpyxl import load_workbook

from adapters.excel.reader import _first, _header, _is_summary_label, _select_prior_sheets
from domain.models import Status
from domain.normalization import normalize_code, normalize_text
from domain.period import AccountingPeriod


_PRIOR_HEADERS=[
    ("거래처코드","거래처 코드","코드"),
    ("거래처명","거래처","업체명"),
    ("적요","내역","내용"),
    ("금액","미지급금","잔액"),
]


def _amount(value):
    value=Decimal(value)
    return int(value) if value == value.to_integral_value() else float(value)


def _change_heading(sheet, header_row, previous, current):
    """Change only recognized previous-month labels above the table header."""
    old_end=date(previous.year,previous.month,calendar.monthrange(previous.year,previous.month)[1])
    new_end=date(current.year,current.month,calendar.monthrange(current.year,current.month)[1])
    for row in sheet.iter_rows(min_row=1,max_row=header_row-1):
        for cell in row:
            value=cell.value
            if isinstance(value,datetime) and value.date()==old_end:
                cell.value=datetime.combine(new_end,value.time())
            elif isinstance(value,date) and value==old_end:
                cell.value=new_end
            elif isinstance(value,str) and not value.startswith("="):
                updated=value
                for sep in ("-", ".", "/"):
                    old_date=f"{previous.year}{sep}{previous.month:02d}{sep}{old_end.day:02d}"
                    new_date=f"{current.year}{sep}{current.month:02d}{sep}{new_end.day:02d}"
                    updated=updated.replace(old_date,new_date)
                for year_fmt in (lambda p: str(p.year),lambda p: f"{p.year%100:02d}"):
                    for pattern in (
                        r"(?<!\d)"+year_fmt(previous)+r"\s*년\s*"+str(previous.month)+r"\s*월",
                        r"(?<!\d)"+year_fmt(previous)+r"\."+f"{previous.month:02d}"+r"(?!\d)",
                    ):
                        updated=re.sub(pattern,lambda m: (
                            f"{year_fmt(current)}년 {current.month}월"
                            if "년" in m.group() else f"{year_fmt(current)}.{current.month:02d}"
                        ),updated)
                cell.value=updated


def _template_date(value, original):
    """Use the original date cell's storage type and visual presentation."""
    if value in ("", None):
        return value
    if isinstance(original,(datetime,date)):
        if isinstance(value,(datetime,date)):
            return value
        try:
            return date.fromisoformat(str(value).replace("/","-").replace(".","-"))
        except ValueError as exc:
            raise ValueError(f"원본 양식의 날짜 형식으로 해석할 수 없는 값: {value}") from exc
    if isinstance(value,(datetime,date)):
        value=value.strftime("%Y-%m-%d")
    text=str(value)
    original_text=str(original or "")
    if len(original_text)>=10:
        if original_text[4:5]=="/" and original_text[7:8]=="/":
            return text.replace("-","/")
        if original_text[4:5]=="." and original_text[7:8]==".":
            return text.replace("-",".")
    return text


def _template_shapes(sheet, header, fields):
    """Identify a familiar flat or group-subtotal template without guessing new columns."""
    detail=None
    subtotal=None
    grouped=False
    code, name, desc, amount=(fields[k] for k in ("code","name","desc","amount"))
    for row in sheet.iter_rows(min_row=header+1):
        if (row[desc-1].value is not None and row[amount-1].value is not None
                and not _is_summary_label(row[desc-1].value)):
            if detail is None:
                detail=row
            if not row[code-1].value and not row[name-1].value:
                grouped=True
        if any(_is_summary_label(row[i-1].value) for i in (fields.get("date"),desc,name) if i):
            if row[code-1].value and row[name-1].value and subtotal is None:
                subtotal=row
    if detail is None:
        raise ValueError("전월 명세서에서 양식으로 사용할 상세 행을 찾지 못했습니다.")
    if grouped and subtotal is None:
        raise ValueError("거래처코드가 소계에만 있는 양식인데 유효한 소계 행을 찾지 못했습니다.")
    def styles(row):
        return ([copy(c._style) for c in row],
                copy(sheet.row_dimensions[row[0].row]))
    return (grouped,styles(detail),styles(subtotal) if subtotal else styles(detail),
            ("date" if fields["date"] and subtotal
             and _is_summary_label(subtotal[fields["date"]-1].value) else "desc"),
            detail[fields["date"]-1].value if fields["date"] else None)



def _status_column(sheet, header_row):
    """Place the status immediately after the rightmost actual header.
    
    A previous generated sheet may already have one. If the user appended
    another column later, move just our status values to the new right edge;
    the other columns and their values are never shifted or overwritten.
    """
    values=[(cell.column, str(cell.value).strip()) for cell in sheet[header_row]
            if cell.value is not None]
    existing=[col for col,txt in values if txt=="처리상태"]
    if len(existing)>1:
        raise ValueError("처리상태 열이 여러 개입니다. 명세서 헤더를 확인해주세요.")
    right=max((col for col,txt in values if txt!="처리상태"),default=0)
    # Also protect data columns whose headings are blank: the next generated
    # status must not overwrite the user's content.
    for cell in sheet._cells.values():
        if cell.row>=header_row and cell.value is not None and cell.column not in existing:
            right=max(right,cell.column)
    if not right:
        raise ValueError("명세서 표의 마지막 열을 식별하지 못했습니다.")
    destination=max(right+1,existing[0] if existing else 0)
    # If the status header is to the left of a newly added user column, move
    # the generated status to the right, leaving all business columns intact.
    if existing and existing[0]<right:
        destination=right+1
        for row in sheet.iter_rows(min_row=header_row):
            row[existing[0]-1].value=None
    return destination


def _status_for_opening(result):
    if result is None:
        return "확인 필요"
    if result.status==Status.MATCHED:
        return ("확인 필요" if result.rule=="CODE_AMOUNT_UNIQUE_WITH_NOTE"
                else "대사 일치")
    if result.status==Status.UNPAID:
        return "전월 이월"
    if result.status==Status.INPUT_TYPO_SUSPECT:
        return "오타 의심"
    if result.status==Status.SIGNED_NET_AUTO:
        return "순잔액 대사"
    return "확인 필요"


def write_month_end_statement(path, prior_paths, results, new_items, issues, period:AccountingPeriod,
                              standalone_debits=(), standalone_debit_sources=None,
                              journal_items=()):
    """Write a *draft*, never a final approved statement. No source is mutated."""
    if not prior_paths:
        raise ValueError("전월 명세서 양식 파일이 필요합니다.")
    out=Path(path).resolve()
    if any(out==Path(p).resolve() for p in prior_paths):
        raise ValueError("원본 명세서에는 덮어쓸 수 없습니다.")
    template_path=Path(prior_paths[0])
    if template_path.suffix.lower()!=".xlsx":
        raise ValueError("원본 양식 보존은 .xlsx 명세서에서만 지원합니다. .xls/.xlsm은 Excel에서 .xlsx로 복사 저장해 주세요.")

    wb=load_workbook(template_path)
    try:
        selected=_select_prior_sheets(wb,_PRIOR_HEADERS,period.previous())
        if len(selected)!=1:
            raise ValueError("당월 양식으로 복사할 전월 명세서 시트를 하나만 식별해야 합니다.")
        original,header,headers=selected[0]
        target_name=f"{period.year%100:02d}.{period.month:02d}"
        if any(ws.title==target_name for ws in wb.worksheets):
            raise ValueError(f"양식 파일에 당월 시트({target_name})가 이미 있습니다. 중복 생성을 중단합니다.")
        fields={
            "code":_first(headers,"거래처코드","거래처 코드","코드"),
            "name":_first(headers,"거래처명","거래처","업체명"),
            "date":_first(headers,"날짜","일자","기표일자"),
            "desc":_first(headers,"적요","내역","내용"),
            "amount":_first(headers,"금액","미지급금","잔액"),
        }
        if not all(fields[k] for k in ("code","name","desc","amount")):
            raise ValueError("거래처코드/거래처명/적요/금액 열을 안전하게 식별하지 못했습니다.")
        grouped,detail_style,subtotal_style,subtotal_label,sample_date=_template_shapes(original,header,fields)

        # Retain the original worksheet object, not a worksheet copy:
        # openpyxl's copy_worksheet loses drawings and some sheet properties.
        # Only data cell values and the period heading are changed in the output.
        ws=original
        if ws.protection.sheet:
            raise ValueError("잠금 설정된 명세서 시트는 양식을 유지한 채 자동 편집할 수 없습니다.")
        if ws._images or ws._charts or ws._pivots:
            raise ValueError("이미지·차트·피벗이 포함된 양식은 보존 여부를 확인할 수 없어 자동 생성을 중단합니다.")
        for region in ws.merged_cells.ranges:
            if region.max_row>header:
                raise ValueError("명세서 본문에 병합 셀이 있어 양식 변경 없이 자동 생성할 수 없습니다.")
        if ws.tables:
            raise ValueError("엑셀 표(Table)로 지정된 명세서는 표 범위 보존을 위해 자동 생성을 중단합니다.")
        # The selected month is replaced *only in the newly saved workbook*.
        # Keep all other workbook sheets, their styles, print settings and layouts.
        ws.title=target_name
        last_template_row=ws.max_row
        # Reject unrecognized content instead of deleting a footer or other data.
        for row in ws.iter_rows(min_row=header+1,max_row=last_template_row):
            values=[cell.value for cell in row]
            if not any(value is not None for value in values):
                continue
            if any(_is_summary_label(values[i-1]) for i in (fields["date"],fields["desc"],fields["name"]) if i):
                continue
            if values[fields["desc"]-1] is not None and values[fields["amount"]-1] is not None:
                continue
            raise ValueError("전월 명세서 본문에 일반 내역·소계 외의 행이 있습니다. 원본 양식을 임의로 삭제하지 않도록 중단합니다.")
        # Capture the *original* row archetypes before physically deleting rows.
        # Detail and subtotal rows must retain different formatting (yellow subtotal
        # background, borders, fonts, number format), including when row counts change.
        # The named sheet is kept: only the original data rows are rebuilt.
        _change_heading(ws,header,period.previous(),period)

        # Preserve *all* original detail lines (including negative rows).
        # The previous statement is a running record; matching a debit by
        # amount does not prove which invoice it settles. No auto-deletion.
        status_column=_status_column(ws,header)
        original_result_rows={
            r.prior.source.row:r for r in results
            if r.prior.source.sheet==ws.title or
               r.prior.source.sheet==original.title or
               r.prior.source.file_name==template_path.name
        }
        net_vendors={
            normalize_code(r.prior.vendor_code):r for r in results
            if r.status in (Status.SIGNED_NET_AUTO,Status.SIGNED_OPENING_REVIEW)
        }
        opening=[]
        pending=[]
        business_columns={col for col in fields.values() if col}
        business_columns.add(status_column)
        subtotal_extra={}
        for idx in range(header+1,last_template_row+1):
            values=[cell.value for cell in ws[idx]]
            if not any(v is not None for v in values):
                continue
            summary=any(_is_summary_label(values[col-1]) for col in
                        (fields["date"],fields["desc"],fields["name"]) if col)
            if grouped and summary:
                code=values[fields["code"]-1]
                name=values[fields["name"]-1]
                if not code or not name:
                    raise ValueError("거래처 소계 행에 코드 또는 거래처명이 없습니다.")
                for source_row,desc,amount,when,extras in pending:
                    opening.append((str(code),str(name),desc,amount,when,source_row,extras))
                for col,val in enumerate(values,1):
                    if col not in business_columns and val is not None:
                        if isinstance(val,str) and val.startswith("="):
                            raise ValueError("추가 열의 소계 수식을 자동 이동하면 참조가 바뀔 수 있어 중단합니다.")
                        subtotal_extra[(str(code),str(name),col)]=val
                pending=[]
                continue
            desc=values[fields["desc"]-1]
            amount=values[fields["amount"]-1]
            if desc is None or amount is None:
                raise ValueError("원본 명세서 상세 행의 적요 또는 금액이 비어 있습니다.")
            if isinstance(amount,bool):
                raise ValueError("명세서 상세 금액은 숫자여야 합니다.")
            try:
                amount=Decimal(str(amount))
            except (ValueError,ArithmeticError) as exc:
                raise ValueError(f"명세서 {idx}행 금액을 안전하게 읽을 수 없습니다.") from exc
            extras={col:val for col,val in enumerate(values,1)
                    if col not in business_columns and val is not None}
            entry=(idx,str(desc),amount,
                   values[fields["date"]-1] if fields["date"] else None,extras)
            if grouped:
                pending.append(entry)
            else:
                code=values[fields["code"]-1]
                name=values[fields["name"]-1]
                if not code or not name:
                    raise ValueError(f"명세서 {idx}행의 거래처가 비어 있습니다.")
                opening.append((str(code),str(name),entry[1],entry[2],entry[3],idx,extras))
        if pending:
            raise ValueError("거래처 소계가 없는 상세 행이 있어 자동 생성할 수 없습니다.")

        records=[]
        audits=[]
        offsets=[]
        reviews=[]
        vendor_order={}
        for code,name,desc,amount,when,source_row,extras in opening:
            vendor=(code,name)
            vendor_order.setdefault(vendor,len(vendor_order))
            result=original_result_rows.get(source_row)
            if result is None:
                result=net_vendors.get(normalize_code(code))
            status=_status_for_opening(result)
            records.append((code,name,desc,amount,when,status,source_row,extras))
            audits.append((status,code,name,desc,_amount(amount),
                           "전월 명세서 원본 행 유지",template_path.name,source_row))
            if status in ("확인 필요","오타 의심"):
                reason=(result.reason if result else "전월 원본 상세 행과 대사결과 연결 확인 필요")
                reviews.append((status,code,name,desc,_amount(amount),
                                reason,template_path.name,source_row))

        # Each accepted period/account-code Douzone RAW line is appended once.
        # The status is informational, never a declaration of payment approval.
        source_debits=standalone_debit_sources or {}
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
        raw=[]
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
            vendor=(line.vendor_code,line.vendor_name)
            vendor_order.setdefault(vendor,len(vendor_order))
            records.append((line.vendor_code,line.vendor_name,line.description,
                            amount,line.date,status,line.row_number,{}))
            audits.append((status,line.vendor_code,line.vendor_name,line.description,
                           _amount(amount),reason,
                           source_debits.get(id(line),"더존 Raw"),line.row_number))
            if status in ("확인 필요","오타 의심","원장 단독"):
                reviews.append((status,line.vendor_code,line.vendor_name,line.description,
                                _amount(amount),reason,
                                source_debits.get(id(line),"더존 Raw"),line.row_number))

        for issue in issues:
            reviews.append(("입력 데이터 확인","","",str(issue.raw_value),"",
                            issue.reason,issue.source,issue.row))
        records.sort(key=lambda record:vendor_order[(record[0],record[1])])
        # Count the exact number of physical rows needed, including subtotal
        # rows. Remove obsolete physical rows rather than leaving blank gaps.
        vendor_keys={(rec[0],rec[1]) for rec in records}
        required_rows=len(records)+(len(vendor_keys) if grouped else 0)
        if last_template_row>header:
            ws.delete_rows(header+1,last_template_row-header)
            for idx in list(ws.row_dimensions):
                if idx>header:
                    del ws.row_dimensions[idx]
        if required_rows:
            ws.insert_rows(header+1,amount=required_rows)

        # Keep the original print-range columns; resize only the bottom row
        # when the number of data rows changes.
        old_print_area=str(ws.print_area or "")
        if old_print_area and required_rows != last_template_row-header:
            from openpyxl.utils.cell import range_boundaries, get_column_letter
            try:
                print_bounds=old_print_area.split("!")[-1].replace("$","")
                min_col,min_row,max_col,max_row=range_boundaries(print_bounds)
            except (TypeError,ValueError):
                raise ValueError("인쇄영역이 복잡한 양식은 범위를 보존할 수 없어 자동 생성을 중단합니다.")
            if max_row != last_template_row:
                raise ValueError("전월 명세서 인쇄영역이 상세 내역 외 행을 포함합니다. 양식 보존을 위해 중단합니다.")
            ws.print_area=(
                f"{get_column_letter(min_col)}{min_row}:"
                f"{get_column_letter(max_col)}{header+required_rows}"
            )

        # Header lives exactly one column to the right of the original table.
        header_status=ws.cell(header,status_column)
        header_status._style=copy(ws.cell(header,status_column-1)._style)
        header_status.value="처리상태"
        ws.column_dimensions[header_status.column_letter].width=17

        row_no=header+1
        def write_row(values,styles,dimension,status):
            nonlocal row_no
            # _style captures font, border, fill, alignment, number format
            # and protection. Match the original row's height and visibility too.
            for column,style in enumerate(styles,1):
                ws.cell(row_no,column)._style=copy(style)
            row_dim=copy(dimension)
            row_dim.index=row_no
            ws.row_dimensions[row_no]=row_dim
            for column,value in values.items():
                ws.cell(row_no,column).value=value
            state_cell=ws.cell(row_no,status_column)
            if status_column>len(styles):
                state_cell._style=copy(ws.cell(row_no,status_column-1)._style)
            if status:
                from openpyxl.styles import Alignment,PatternFill
                state_cell.value=status
                state_cell.alignment=Alignment(horizontal="center",vertical="center")
                state_cell.number_format="General"
                colors={
                    "대사 일치":"E2F0D9",
                    "전월 이월":"E9EDF4",
                    "순잔액 대사":"E2F0D9",
                    "당월 발생":"EEE3FF",
                    "당월 지급":"E6F2FF",
                    "원장 단독":"FFF2CC",
                    "확인 필요":"FFE4CC",
                    "오타 의심":"FADBD8",
                }
                state_cell.fill=PatternFill(fill_type="solid",fgColor=colors.get(status,"FFFFFF"))
            row_no+=1

        def finish_subtotal(vendor,start,total):
            if not grouped:
                return
            column_letter=ws.cell(row_no,fields["amount"]).column_letter
            values={
                fields["code"]:vendor[0],fields["name"]:vendor[1],
                fields["amount"]:f"=SUM({column_letter}{start}:{column_letter}{row_no-1})",
            }
            values[fields["date"] if subtotal_label=="date" else fields["desc"]]="소계"
            for (code,name,column),value in subtotal_extra.items():
                if (code,name)==vendor:
                    values[column]=value
            write_row(values,*subtotal_style,status="소계")
        vendor=None
        start=None
        total=Decimal(0)
        for rec in records:
            code,name,desc,amount,row_date,status,source_row,extras=rec
            current=(code,name)
            if vendor is not None and current!=vendor:
                finish_subtotal(vendor,start,total)
                start=None
                total=Decimal(0)
            if start is None:
                start=row_no
            if row_date is None:
                row_date=f"{period.year}-{period.month:02d}-{calendar.monthrange(period.year,period.month)[1]}"
            values={fields["desc"]:desc,fields["amount"]:_amount(amount)}
            values.update(extras)
            if fields["date"]:
                values[fields["date"]]=_template_date(row_date,sample_date)
            if not grouped:
                values[fields["code"]]=code
                values[fields["name"]]=name
            write_row(values,*detail_style,status=status)
            total+=Decimal(amount)
            vendor=current
        if vendor is not None:
            finish_subtotal(vendor,start,total)
        if row_no != header+required_rows+1:
            raise AssertionError("당월 명세서 행 수와 실제 생성 행 수가 일치하지 않습니다.")

        # Fully offset details are absent from the main statement but every
        # source/target pair remains traceable in this removable auxiliary tab.
        offset_sheet=wb.create_sheet("대사내역")
        offset_sheet.append([
            "구분","거래처코드","거래처명","원본 적요","발생 금액","발생 일자",
            "원본파일","원본행","더존 RAW 적요","더존 기표일자",
            "더존 차변","상계 확정 여부","더존 RAW 행",
        ])
        for row in offsets:
            offset_sheet.append(row)
        offset_sheet.freeze_panes="A2"
        for column,width in {"A":24,"B":17,"C":26,"D":37,"E":18,"F":18,
                             "G":35,"H":12,"I":37,"J":18,"K":20,"L":20,"M":15}.items():
            offset_sheet.column_dimensions[column].width=width
        for row in offset_sheet.iter_rows(min_row=2):
            for idx in (4,10,11):
                row[idx].number_format="#,##0;[Red](#,##0);-"

        # No extra columns or hard-to-erase tags in the original template sheet.
        summary=wb.create_sheet("검토필요")
        summary.append(["검토 상태","거래처코드","거래처명","적요·원본값","금액","사유","원본파일","원본행"])
        summary["J1"]="MONTH_END_DRAFT_REVIEW"
        for record in reviews:
            summary.append(record)
        summary.freeze_panes="A2"
        changes=wb.create_sheet("변경내역")
        changes.append(["구분","거래처코드","거래처명","명세 내용","금액","근거","원본파일","원본행"])
        for record in audits:
            changes.append(record)
        changes.freeze_panes="A2"
        changes["J1"]="※ 검토용 초안: 발생·지급은 모두 보존하며 대사 일치는 실제 상계 확정과 다릅니다."
        changes["J2"]="대사내역·검토필요·변경내역은 검토 후 삭제할 수 있습니다. 검토 보류건 확정 전에는 다음 달 대사에 사용하지 마세요."
        changes["J3"]=f"검토 필요 {len(reviews)}건 · RAW 전표 {len(seen)}건 · 일치 후보 {len(offsets)}건"
        for extra in (summary,changes):
            extra.column_dimensions["A"].width=23
            extra.column_dimensions["B"].width=16
            extra.column_dimensions["C"].width=25
            extra.column_dimensions["D"].width=36
            extra.column_dimensions["E"].width=18
            extra.column_dimensions["F"].width=65
            extra.column_dimensions["G"].width=32
            extra.column_dimensions["H"].width=12
        changes.column_dimensions["J"].width=75

        # Let Excel recalculate all copied per-vendor subtotal formulas on open.
        from openpyxl.workbook.properties import CalcProperties
        wb.calculation=CalcProperties(calcMode="auto",fullCalcOnLoad=True,forceFullCalc=True)
        wb.active=wb.worksheets.index(ws)
        wb.save(out)
    finally:
        wb.close()
