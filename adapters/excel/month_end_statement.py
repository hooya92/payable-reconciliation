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

        # Keep review opening balances unchanged; *never* assume suspected payments
        # or corrections are approved. Unverifiable new credits stay off the draft.
        records=[]
        audits=[]
        offsets=[]
        reviews=[]
        standalone_debit_sources=standalone_debit_sources or {}
        for result in results:
            p=result.prior
            status=result.status
            if status==Status.MATCHED:
                if result.rule=="CODE_AMOUNT_UNIQUE_WITH_NOTE" and result.journal is not None:
                    # Matching amount alone is enough for reconciliation status,
                    # but a changed invoice description/name should not trigger
                    # irreversible removal from a human-edited month-end draft.
                    records.append((p.vendor_code,p.vendor_name,None,None,p.description,p.amount,
                                    p.source.file_name,"검토보류",p.source.row,p.date))
                    reviews.append(("대사 참고사항 확인",p.vendor_code,p.vendor_name,p.description,
                                    _amount(p.amount),result.reason,p.source.file_name,p.source.row))
                    audits.append(("검토 전 보존",p.vendor_code,p.vendor_name,p.description,
                                   _amount(p.amount),result.reason,p.source.file_name,p.source.row))
                    continue
                if result.journal is not None:
                    j=result.journal
                    offsets.append(("전월 전액 상계",p.vendor_code,p.vendor_name,p.description,
                                    _amount(p.amount),p.date,p.source.file_name,p.source.row,
                                    j.description,j.date,_amount(j.debit),0,j.row_number))
                    audits.append(("전액 상계·본문 제외",p.vendor_code,p.vendor_name,p.description,
                                   _amount(p.amount),result.reason,p.source.file_name,p.source.row))
                else:
                    audits.append(("지급 완료 제외",p.vendor_code,p.vendor_name,p.description,
                                   _amount(p.amount),result.reason,p.source.file_name,p.source.row))
                # The positive opening and its equal verified debit cancel:
                # neither should remain in the outstanding balance statement.
                continue
            if status==Status.SIGNED_NET_AUTO:
                if result.closing_balance is None:
                    reviews.append(("거래처 순잔액 미확정",p.vendor_code,p.vendor_name,p.description,_amount(p.amount),result.reason,p.source.file_name,p.source.row))
                    continue
                if result.closing_balance<=0:
                    audits.append(("순잔액 0원 제외",p.vendor_code,p.vendor_name,p.description,
                                   _amount(p.amount),result.reason,p.source.file_name,p.source.row))
                    continue
                description="거래처 미지급 순잔액 (개별 청구 건 배분 미확정)"
                records.append((p.vendor_code,p.vendor_name,period.year,period.month,description,
                                result.closing_balance,p.source.file_name,"거래처순잔액",p.source.row))
                audits.append(("거래처 순잔액 반영",p.vendor_code,p.vendor_name,description,
                               _amount(result.closing_balance),result.reason,p.source.file_name,p.source.row))
                continue

            reason=("당월 더존 대응 차변 없음" if status==Status.UNPAID
                    else result.reason or status.value)
            kind="이월" if status==Status.UNPAID else "검토보류"
            records.append((p.vendor_code,p.vendor_name,None,None,p.description,p.amount,
                            p.source.file_name,kind,p.source.row,p.date))
            audits.append((kind,p.vendor_code,p.vendor_name,p.description,_amount(p.amount),
                           reason,p.source.file_name,p.source.row))
            if kind=="검토보류":
                reviews.append((status.value,p.vendor_code,p.vendor_name,p.description,
                                _amount(p.amount),reason,p.source.file_name,p.source.row))
        for item in new_items:
            j=item.journal
            if not item.auto_carry:
                reviews.append(("당월 신규 확인 필요",j.vendor_code,j.vendor_name,j.description,
                                _amount(j.credit),item.reason,"더존 Raw",j.row_number))
                continue
            if item.remaining<j.credit:
                # These offsets were independently validated by new_payables().
                # Audit them even if the remaining balance is zero and no detail
                # row is carried to the closing statement.
                debits=[
                    d for d in journal_items
                    if (normalize_code(d.vendor_code)==normalize_code(j.vendor_code)
                        and normalize_text(d.description)==normalize_text(j.description)
                        and d.debit>0)
                ]
                for debit in debits:
                    offsets.append(("당월 발생·차변 상계",j.vendor_code,j.vendor_name,j.description,
                                    _amount(j.credit),j.date,"더존 Raw",j.row_number,
                                    debit.description,debit.date,_amount(debit.debit),
                                    _amount(item.remaining),debit.row_number))
            if item.remaining<=0:
                audits.append(("당월 발생분 전액 상계",j.vendor_code,j.vendor_name,j.description,
                               _amount(j.credit),item.reason,"더존 Raw",j.row_number))
                continue
            records.append((j.vendor_code,j.vendor_name,None,None,j.description,
                            item.remaining,"더존 Raw","신규",j.row_number,j.date))
            audits.append(("당월 신규 반영",j.vendor_code,j.vendor_name,j.description,
                           _amount(item.remaining),item.reason,"더존 Raw",j.row_number))

        for j in standalone_debits:
            if j.debit<=0:
                continue
            # Preserve the exact Douzone description and date rather than
            # fabricating the purpose of the unmatched payment/adjustment.
            source_name=standalone_debit_sources.get(id(j),"더존 Raw")
            records.append((j.vendor_code,j.vendor_name,None,None,j.description,
                            -j.debit,source_name,"원장단독차변",j.row_number,j.date))
            audits.append(("더존 단독 차변 반영",j.vendor_code,j.vendor_name,j.description,
                           _amount(-j.debit),"전월·당월 대응 발생 없음 / RAW 기재 내역 그대로 반영",
                           source_name,j.row_number))
            # A negative payable balance is not a confirmed ordinary payable.
            # Do not silently approve it as next month's opening balance.
            reviews.append(("음수 순잔액 참고",j.vendor_code,j.vendor_name,j.description,
                            _amount(-j.debit),"RAW 내역은 반영됨. 음수 잔액이므로 다음 달 확정 전 확인",
                            source_name,j.row_number))
        for issue in issues:
            reviews.append(("입력 데이터 확인", "", "", str(issue.raw_value),
                            "",issue.reason,issue.source,issue.row))

        # Preserve the previous statement's vendor order and invoice order.
        # New invoices for an existing vendor join its group; truly new vendors
        # are appended in the order they first appeared in the Douzone RAW.
        vendor_order={}
        for result in results:
            vendor=(result.prior.vendor_code,result.prior.vendor_name)
            vendor_order.setdefault(vendor,len(vendor_order))
        for item in new_items:
            vendor=(item.journal.vendor_code,item.journal.vendor_name)
            vendor_order.setdefault(vendor,len(vendor_order))
        records.sort(key=lambda row:vendor_order[(row[0],row[1])])
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

        row_no=header+1
        def write_row(values,styles,dimension):
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
            write_row(values,*subtotal_style)
        vendor=None
        start=None
        total=Decimal(0)
        for rec in records:
            code,name,_,_,desc,amount,source,kind,source_row,*dates=rec
            current=(code,name)
            if vendor is not None and current!=vendor:
                finish_subtotal(vendor,start,total)
                start=None
                total=Decimal(0)
            if start is None:
                start=row_no
            if len(dates):
                row_date=dates[0]
            else:
                # Grouped signed balance uses the month-end date.
                row_date=f"{period.year}-{period.month:02d}-{calendar.monthrange(period.year,period.month)[1]}"
            values={fields["desc"]:desc,fields["amount"]:_amount(amount)}
            if fields["date"]:
                values[fields["date"]]=_template_date(row_date,sample_date)
            if not grouped:
                values[fields["code"]]=code
                values[fields["name"]]=name
            write_row(values,*detail_style)
            total+=Decimal(amount)
            vendor=current
        if vendor is not None:
            finish_subtotal(vendor,start,total)
        if row_no != header+required_rows+1:
            raise AssertionError("당월 명세서 행 수와 실제 생성 행 수가 일치하지 않습니다.")

        # Fully offset details are absent from the main statement but every
        # source/target pair remains traceable in this removable auxiliary tab.
        offset_sheet=wb.create_sheet("상계내역")
        offset_sheet.append([
            "구분","거래처코드","거래처명","원본 적요","발생 금액","발생 일자",
            "원본파일","원본행","더존 RAW 적요","더존 기표일자",
            "더존 차변(상계)","상계 후 잔액","더존 RAW 행",
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
        changes["J1"]="※ 검토용 초안: 미확정 행은 원금액 유지 / 검토필요 시트 확인 후 확정"
        changes["J2"]="상계내역·검토필요·변경내역은 검토 후 삭제할 수 있습니다. 검토 보류건 확정 전에는 다음 달 대사에 사용하지 마세요."
        changes["J3"]=f"검토 필요 {len(reviews)}건 · 본문 미확정 항목 {sum(rec[7]=='검토보류' for rec in records)}건"
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
