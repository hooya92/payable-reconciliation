"""Create an editable month-end statement using the previous statement as a visual template.

The input workbook is never saved. The main sheet keeps the template columns;
review explanations and change history live on removable companion sheets.
"""
from __future__ import annotations

from collections import defaultdict
from copy import copy
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
import calendar
import re

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from adapters.excel.reader import _first, _header, _is_summary_label, _select_prior_sheets
from domain.models import Status
from domain.period import AccountingPeriod


_PRIOR_HEADERS=[
    ("거래처코드","거래처 코드","코드"),
    ("거래처명","거래처","업체명"),
    ("적요","내역","내용"),
    ("금액","미지급금","잔액"),
]
_REVIEW=PatternFill(fill_type="solid",fgColor="FFF2CC")
_NEW=PatternFill(fill_type="solid",fgColor="EDE7F6")


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
        return [copy(c._style) for c in row], sheet.row_dimensions[row[0].row].height
    return grouped,styles(detail),styles(subtotal) if subtotal else styles(detail),("date" if fields["date"] and subtotal and _is_summary_label(subtotal[fields["date"]-1].value) else "desc")


def write_month_end_statement(path, prior_paths, results, new_items, issues, period:AccountingPeriod):
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
        grouped,detail_style,subtotal_style,subtotal_label=_template_shapes(original,header,fields)

        # Work only on a copy, so template content/styles stay unmodified.
        ws=wb.copy_worksheet(original)
        ws.title=target_name
        # The openpyxl worksheet copier does not preserve every merged range.
        for region in original.merged_cells.ranges:
            if region.max_row<=header and str(region) not in ws.merged_cells:
                ws.merge_cells(str(region))
        ws.sheet_view.showGridLines=False
        for region in list(ws.merged_cells.ranges):
            if region.max_row>header:
                if region.min_row<=header:
                    raise ValueError("양식의 병합 셀이 헤더와 상세 행에 걸쳐 있어 자동 재구성이 어렵습니다.")
                ws.unmerge_cells(str(region))
        if ws.max_row>header:
            ws.delete_rows(header+1,ws.max_row-header)
        _change_heading(ws,header,period.previous(),period)
        ws.auto_filter.ref=None

        # Keep review opening balances unchanged; *never* assume suspected payments
        # or corrections are approved. Unverifiable new credits stay off the draft.
        records=[]
        audits=[]
        reviews=[]
        for result in results:
            p=result.prior
            status=result.status
            if status==Status.MATCHED:
                audits.append(("지급 완료 제외",p.vendor_code,p.vendor_name,p.description,
                               _amount(p.amount),"더존 차변 일치",p.source.file_name,p.source.row))
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
            if not item.auto_carry or item.remaining<=0:
                reviews.append(("당월 신규 확인 필요",j.vendor_code,j.vendor_name,j.description,
                                _amount(j.credit),item.reason,"더존 Raw",j.row_number))
                continue
            records.append((j.vendor_code,j.vendor_name,None,None,j.description,
                            item.remaining,"더존 Raw","신규",j.row_number,j.date))
            audits.append(("당월 신규 반영",j.vendor_code,j.vendor_name,j.description,
                           _amount(item.remaining),item.reason,"더존 Raw",j.row_number))
        for issue in issues:
            reviews.append(("입력 데이터 확인", "", "", str(issue.raw_value),
                            "",issue.reason,issue.source,issue.row))

        # One stable vendor group makes it possible to delete any detail row in Excel.
        records.sort(key=lambda x:(str(x[1]),str(x[0]),str(x[4]),str(x[8])))
        row_no=header+1
        def write_row(values,styles,height,highlight=None):
            nonlocal row_no
            for column in range(1,max(len(styles),max(value for value in fields.values() if value))+1):
                cell=ws.cell(row_no,column)
                if column<=len(styles):
                    cell._style=copy(styles[column-1])
                if column in values:
                    cell.value=values[column]
            if highlight:
                ws.cell(row_no,fields["amount"]).fill=copy(highlight)
            if height is not None:
                ws.row_dimensions[row_no].height=height
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
                values[fields["date"]]=row_date
            if not grouped:
                values[fields["code"]]=code
                values[fields["name"]]=name
            write_row(values,*detail_style,highlight=_REVIEW if kind=="검토보류" else _NEW if kind=="신규" else None)
            total+=Decimal(amount)
            vendor=current
        if vendor is not None:
            finish_subtotal(vendor,start,total)

        # No extra columns or hard-to-erase tags in the original template sheet.
        ws.sheet_properties.tabColor="FFD966" if reviews else "70AD47"
        summary=wb.create_sheet("검토필요")
        summary.append(["검토 상태","거래처코드","거래처명","적요·원본값","금액","사유","원본파일","원본행"])
        for record in reviews:
            summary.append(record)
        summary.freeze_panes="A2"
        changes=wb.create_sheet("변경내역")
        changes.append(["구분","거래처코드","거래처명","적요","금액","근거","원본파일","원본행"])
        for record in audits:
            changes.append(record)
        changes.freeze_panes="A2"
        changes["J1"]="※ 검토용 초안: 미확정 행은 원금액 유지 / 검토필요 시트 확인 후 확정"
        changes["J2"]="검토필요·변경내역 시트는 검토 후 통째로 삭제할 수 있습니다."
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

        for sheet in list(wb.worksheets):
            if sheet not in (ws,summary,changes):
                wb.remove(sheet)
        wb.active=0
        wb.save(out)
    finally:
        wb.close()
