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
from openpyxl.comments import Comment
from openpyxl.utils.cell import get_column_letter, quote_sheetname
from openpyxl.styles.cell_style import StyleArray
from openpyxl.worksheet.hyperlink import Hyperlink
from adapters.excel.text_safe import force_text_formulas, set_text

from adapters.excel.reader import (
    _first, _header, _is_blank, _is_summary_label, _select_prior_sheets, _statement_row_kinds,
)
from domain.models import Status, DataQuality
from domain.normalization import parse_amount, parse_literal_amount_formula
from application.month_end_plan import build_month_end_plan
from adapters.excel.month_end_status import (
    find_status_column, is_generated_status, link_note_text, prepare_status_header, write_status,
)
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


class _AsText(str):
    """A value that must be written as text even when it starts with "=" (a copied formula)."""


def _stale_status_columns(sheet, header, business_columns):
    """Columns that only hold status words this program wrote in an earlier month.

    Such a column (no header, every value one of our own status labels) is last
    month's output, not the owner's notes, and must not be copied into the new
    month: it would show last month's status next to this month's rows.
    A column with any free-form text or its own header is never touched.
    """
    found={}
    for row in sheet.iter_rows(min_row=header+1):
        for cell in row:
            if cell.column in business_columns or _is_blank(cell.value):
                continue
            found.setdefault(cell.column,[]).append(cell.value)
    return {
        column for column,values in found.items()
        if _is_blank(sheet.cell(header,column).value)
        and all(is_generated_status(value) for value in values)
    }


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


def _template_shapes(sheet, header, fields, auxiliary_rows=()):
    """Identify a familiar flat or group-subtotal template without guessing new columns."""
    detail=None
    subtotal=None
    grouped=False
    code, name, desc, amount=(fields[k] for k in ("code","name","desc","amount"))
    for row in sheet.iter_rows(min_row=header+1):
        if row[0].row in auxiliary_rows:
            continue
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
        raise ValueError("전월 명세서에서 양식으로 사용할 상세 행을 찾지 못했습니다. 적요와 금액이 모두 있는 거래 행이 하나 이상 필요합니다.")
    def styles(row):
        return ([copy(c._style) for c in row],
                copy(sheet.row_dimensions[row[0].row]))
    return (grouped,styles(detail),styles(subtotal) if subtotal else styles(detail),
            ("date" if fields["date"] and subtotal
             and _is_summary_label(subtotal[fields["date"]-1].value) else "desc"),
            detail[fields["date"]-1].value if fields["date"] else None,
            grouped and subtotal is None)



def write_month_end_statement(path, prior_paths, results, new_items, issues, period:AccountingPeriod,
                              standalone_debits=(), standalone_debit_sources=None,
                              journal_items=(), journal_sources=None, source_paths=()):
    """Write a *draft*, never a final approved statement. No source is mutated."""
    if not prior_paths:
        raise ValueError("전월 명세서 양식 파일이 필요합니다.")
    if len(prior_paths)!=1:
        raise ValueError(
            "당월 명세서 자동 생성은 현재 전월 양식 1개 파일만 지원합니다. "
            "여러 담당자 양식 중 첫 파일만 사용하면 다른 거래처가 누락될 수 있어 생성을 중단합니다."
        )
    out=Path(path).resolve()
    if any(out==Path(p).resolve() for p in (*prior_paths, *source_paths)):
        raise ValueError("원본 Excel에는 덮어쓸 수 없습니다.")
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
            raise ValueError(f"양식 파일에 당월 시트({target_name})가 이미 있습니다. 그 시트를 삭제하거나 이름을 바꾼 뒤 다시 실행해 주세요.")
        fields={
            "code":_first(headers,"거래처코드","거래처 코드","코드"),
            "name":_first(headers,"거래처명","거래처","업체명"),
            "date":_first(headers,"날짜","일자","기표일자"),
            "desc":_first(headers,"적요","내역","내용"),
            "amount":_first(headers,"금액","미지급금","잔액"),
        }
        if not all(fields[k] for k in ("code","name","desc","amount")):
            raise ValueError("거래처코드/거래처명/적요/금액 열을 안전하게 식별하지 못했습니다.")
        auxiliary_columns={"vendor_code":fields["code"],"vendor_name":fields["name"],
                           "date":fields["date"],"description":fields["desc"],
                           "amount":fields["amount"]}
        auxiliary_rows=_statement_row_kinds(
            [(row[0].row,[cell.value for cell in row])
             for row in original.iter_rows(min_row=header+1)],
            auxiliary_columns,
        )
        grouped,detail_style,subtotal_style,subtotal_label,sample_date,no_subtotal_sample=_template_shapes(
            original,header,fields,auxiliary_rows
        )
        # Rows that cannot be read are never a reason to stop: they are kept as they were,
        # left out of the reconciliation, and listed so the user can look at them.
        held={}
        notices=[]
        unmerged_rows={}
        def hold(idx,reason):
            if idx in unmerged_rows:
                # The value may have been hidden behind a merge, which is why it reads as empty.
                reason=f"{reason} (병합 셀 {unmerged_rows[idx]}에 가려진 값일 수 있음)"
            held.setdefault(idx,reason)
        if no_subtotal_sample:
            notices.append("전월 명세서에서 거래처 소계 행을 찾지 못해, 새로 만드는 소계 행은 상세 행과 같은 서식으로 만들었습니다. "
                           "소계 행은 날짜·적요·거래처명 열 중 한 곳에 '소계'라고 적고 같은 행에 거래처코드·거래처명이 있어야 인식됩니다.")

        # Retain the original worksheet object, not a worksheet copy:
        # openpyxl's copy_worksheet loses drawings and some sheet properties.
        # Only data cell values and the period heading are changed in the output.
        ws=original
        if ws.protection.sheet:
            raise ValueError(f"'{ws.title}' 시트가 보호(잠금)되어 있습니다. Excel의 검토 > 시트 보호 해제 후 다시 저장해 주세요.")
        if ws._images or ws._charts or ws._pivots:
            raise ValueError(f"'{ws.title}' 시트에 이미지·차트·피벗이 있어 양식을 안전하게 보존할 수 없습니다. 해당 개체를 제거한 복사본으로 다시 실행해 주세요.")
        auxiliary_merges=[]
        for region in list(ws.merged_cells.ranges):
            if region.max_row>header:
                if all(idx in auxiliary_rows for idx in range(region.min_row,region.max_row+1)):
                    auxiliary_merges.append(copy(region))
                else:
                    # Rows are rebuilt, so a merge across transaction rows cannot follow them.
                    # Only the new file loses the merge; the original is never touched.
                    ws.unmerge_cells(region.coord)
                    for merged_row in range(region.min_row,region.max_row+1):
                        unmerged_rows[merged_row]=region.coord
                    notices.append(f"명세서 본문의 병합 셀({region.coord})은 행이 다시 만들어지는 과정에서 유지할 수 없어 "
                                   "새 파일에서만 병합을 해제했습니다. 병합된 칸에 가려져 있던 값은 읽을 수 없습니다.")
        if ws.tables:
            raise ValueError("명세서가 엑셀 '표(Table)'로 지정되어 있어 행 수가 바뀌면 표 범위가 깨집니다. 표 디자인 > 범위로 변환 후 다시 실행해 주세요.")
        # The selected month is replaced *only in the newly saved workbook*.
        # Keep all other workbook sheets, their styles, print settings and layouts.
        source_sheet=ws.title
        ws.title=target_name
        last_template_row=ws.max_row
        unclear_rows=set()
        # Auxiliary rows are excluded from accounting and restored below the table.
        for row in ws.iter_rows(min_row=header+1,max_row=last_template_row):
            if row[0].row in auxiliary_rows:
                continue
            values=[cell.value for cell in row]
            if all(_is_blank(value) for value in values):
                continue
            if any(_is_summary_label(values[i-1]) for i in (fields["date"],fields["desc"],fields["name"]) if i):
                continue
            if not _is_blank(values[fields["desc"]-1]) and not _is_blank(values[fields["amount"]-1]):
                continue
            unclear_rows.add(row[0].row)
            hold(row[0].row,"적요 또는 금액이 비어 있어 거래 행인지 소계·보조 행인지 구분할 수 없음")
        # Capture the *original* row archetypes before physically deleting rows.
        # Detail and subtotal rows must retain different formatting (yellow subtotal
        # background, borders, fonts, number format), including when row counts change.
        # The named sheet is kept: only the original data rows are rebuilt.
        _change_heading(ws,header,period.previous(),period)

        # Preserve all original business columns, including free-form notes.
        # Only the optional processed-status column may be newly appended.
        status_column=find_status_column(ws,header,ignore_rows=set(auxiliary_rows))
        opening=[]
        pending=[]
        business_columns={col for col in fields.values() if col}
        business_columns.add(status_column)
        stale_columns=_stale_status_columns(ws,header,business_columns)
        subtotal_extra={}
        subtotal_styles={}
        for idx in range(header+1,last_template_row+1):
            if idx in auxiliary_rows or idx in unclear_rows:
                continue
            values=[cell.value for cell in ws[idx]]
            if all(_is_blank(v) for v in values):
                continue
            summary=any(_is_summary_label(values[col-1]) for col in
                        (fields["date"],fields["desc"],fields["name"]) if col)
            if grouped and summary:
                code=values[fields["code"]-1]
                name=values[fields["name"]-1]
                if not code or not name:
                    hold(idx,"거래처 소계 행에 거래처코드 또는 거래처명이 없음")
                    for entry in pending:
                        hold(entry[0],"소계 행에 거래처코드·거래처명이 없어 거래처를 확정할 수 없음")
                    pending=[]
                    continue
                subtotal_styles[(str(code),str(name))]=(
                    [copy(c._style) for c in ws[idx]],copy(ws.row_dimensions[idx])
                )
                for source_row,desc,amount,when,extras in pending:
                    opening.append((str(code),str(name),desc,amount,when,source_row,extras))
                for col,val in enumerate(values,1):
                    if col not in business_columns and col not in stale_columns and val is not None:
                        if isinstance(val,str) and val.startswith("="):
                            # Moving a formula would change what it points at: carry it as text.
                            val=_AsText(val)
                            notices.append(f"명세서 {idx}행 {get_column_letter(col)}열의 수식은 옮기면 참조가 바뀌므로 텍스트로 옮겼습니다.")
                        subtotal_extra[(str(code),str(name),col)]=val
                pending=[]
                continue
            desc=values[fields["desc"]-1]
            amount=values[fields["amount"]-1]
            if _is_blank(desc) or _is_blank(amount):
                hold(idx,"적요 또는 금액이 비어 있음")
                continue
            if isinstance(amount,bool):
                hold(idx,"금액이 TRUE/FALSE 값임")
                continue
            if isinstance(amount,str) and amount.startswith("="):
                # A literal has no references or stale calculation cache.
                # Other formulas require review rather than guessed evaluation.
                literal=parse_literal_amount_formula(amount)
                if literal is None:
                    hold(idx,"금액에 지원하지 않는 수식이 있음")
                    continue
                amount=literal
            elif isinstance(amount,str):
                parsed=parse_amount(amount, allow_zero_marker=False)
                if parsed.quality==DataQuality.SUSPICIOUS or parsed.value is None:
                    hold(idx,f"금액 형식을 확정할 수 없음 ({amount})")
                    continue
                amount=parsed.value
            try:
                amount=Decimal(str(amount))
            except (ValueError,ArithmeticError):
                hold(idx,f"금액을 안전하게 읽을 수 없음 ({amount!r})")
                continue
            extras={col:val for col,val in enumerate(values,1)
                    if col not in business_columns and col not in stale_columns
                    and val is not None}
            entry=(idx,str(desc),amount,
                   values[fields["date"]-1] if fields["date"] else None,extras)
            if grouped:
                pending.append(entry)
            else:
                code=values[fields["code"]-1]
                name=values[fields["name"]-1]
                if not code or not name:
                    hold(idx,"거래처코드 또는 거래처명이 비어 있음")
                    continue
                opening.append((str(code),str(name),entry[1],entry[2],entry[3],idx,extras))
        for entry in pending:
            hold(entry[0],"이 상세 행 아래에 거래처 소계 행이 없어 거래처를 확정할 수 없음")
        # Unreadable rows stay where the notes stay: kept as they were, outside the accounting.
        auxiliary_rows=dict(sorted({**auxiliary_rows,**{idx:"held" for idx in held}}.items()))
        for idx in sorted(held):
            notices.append(f"명세서 {idx}행: {held[idx]}. 대사에서 제외하고 원래 모습 그대로 두었습니다.")

        # Each retained opening row keeps its own formatting. New RAW rows
        # alone use the existing template's representative detail style.
        opening_styles={
            id(entry[6]):([copy(c._style) for c in ws[entry[5]]],
                          copy(ws.row_dimensions[entry[5]]))
            for entry in opening
        }
        opening_amounts={
            id(entry[6]):ws.cell(entry[5],fields["amount"]).value
            for entry in opening
        }
        plan=build_month_end_plan(
            opening,results,new_items,issues,template_path.name,source_sheet,
            standalone_debits=standalone_debits,
            standalone_debit_sources=standalone_debit_sources,
            journal_items=journal_items,
            journal_sources=journal_sources,
        )
        records,audits,offsets,reviews=(
            plan.records,plan.audits,plan.offsets,plan.reviews
        )
        auxiliary_cells={idx:[copy(cell) for cell in ws[idx]] for idx in auxiliary_rows}
        auxiliary_dimensions={idx:copy(ws.row_dimensions[idx]) for idx in auxiliary_rows}
        cached=None
        if any(cell.data_type=="f" for cells in auxiliary_cells.values() for cell in cells):
            cached_wb=load_workbook(template_path,data_only=True)
            try:
                cached={cell.coordinate:cached_wb[source_sheet][cell.coordinate].value
                        for cells in auxiliary_cells.values() for cell in cells
                        if cell.data_type=="f"}
            finally:
                cached_wb.close()
        for region in auxiliary_merges:
            ws.unmerge_cells(str(region))
        # Count the exact number of physical rows needed, including subtotal
        # rows. Remove obsolete physical rows rather than leaving blank gaps.
        vendor_keys={(rec.vendor_code,rec.vendor_name) for rec in records}
        body_rows=len(records)+(len(vendor_keys) if grouped else 0)
        required_rows=body_rows+len(auxiliary_rows)
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

        prepare_status_header(ws,header,status_column,fields["amount"])

        row_no=header+1
        source_targets={}
        def write_row(values,styles,dimension):
            nonlocal row_no
            # The original cell style carries font, border, fill, number format,
            # alignment and protection. Row properties are copied as well.
            for column,style in enumerate(styles,1):
                ws.cell(row_no,column)._style=(
                    StyleArray() if column in stale_columns else copy(style))
            row_dim=copy(dimension)
            row_dim.index=row_no
            ws.row_dimensions[row_no]=row_dim
            for column,value in values.items():
                # Statement text from the source must stay text, never a formula.
                if column in (fields["code"],fields["name"],fields["desc"]) or isinstance(value,_AsText):
                    set_text(ws.cell(row_no,column),str(value))
                else:
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
            for (code,name,column),value in subtotal_extra.items():
                if (code,name)==vendor:
                    values[column]=value
            write_row(values,*subtotal_styles.get(vendor,subtotal_style))
            write_status(ws,row_no-1,status_column,fields["amount"],"소계",subtotal=True,
                         subtotal_zero=(total==0),note=plan.subtotal_notes.get(vendor))
        vendor=None
        start=None
        total=Decimal(0)
        record_rows={}
        for rec in records:
            current=(rec.vendor_code,rec.vendor_name)
            if vendor is not None and current!=vendor:
                finish_subtotal(vendor,start,total)
                start=None
                total=Decimal(0)
            if start is None:
                start=row_no
            row_date=rec.date
            if row_date is None:
                row_date=f"{period.year}-{period.month:02d}-{calendar.monthrange(period.year,period.month)[1]}"
            values={fields["desc"]:rec.description,fields["amount"]:_amount(rec.amount)}
            values.update(rec.extras)
            original_style=opening_styles.get(id(rec.extras))
            if original_style is not None:
                original_amount=opening_amounts[id(rec.extras)]
                # Text amounts must become numeric in the output so Excel SUM
                # includes them. Keep the original formatting and input file.
                values[fields["amount"]]=(
                    _amount(rec.amount) if isinstance(original_amount,str)
                    and not original_amount.startswith("=") else original_amount
                )
            if fields["date"]:
                values[fields["date"]]=(
                    rec.date if original_style is not None
                    else _template_date(row_date,sample_date)
                )
            if not grouped:
                values[fields["code"]]=rec.vendor_code
                values[fields["name"]]=rec.vendor_name
            source_ref=plan.record_sources.get(id(rec))
            if source_ref is not None:
                source_targets.setdefault(source_ref,[]).append(row_no)
            write_row(values,*(original_style or detail_style))
            record_rows[id(rec)]=row_no-1
            write_status(ws,row_no-1,status_column,fields["amount"],rec.status,
                         reason=plan.status_reasons.get(id(rec),""),
                         text=plan.status_notes.get(id(rec)))
            total+=Decimal(rec.amount)
            vendor=current
        if vendor is not None:
            finish_subtotal(vendor,start,total)
        # Say which rows each row is offset against ("8행과 대차"). Row numbers are final
        # here (nothing is inserted afterwards). Plain text: no link, the rows are close.
        for rec in records:
            note=plan.link_notes.get(id(rec))
            if note is None or id(rec) not in record_rows:
                continue
            partner_rows=[record_rows[p] for p in note.partners if p in record_rows]
            if partner_rows:
                cell=ws.cell(record_rows[id(rec)],status_column)
                text=link_note_text(note,partner_rows)
                cell.value=f"{cell.value}\n{text}" if note.append and cell.value else text
        auxiliary_targets={}
        for source_row,kind in auxiliary_rows.items():
            auxiliary_targets[source_row]=row_no
            cells=auxiliary_cells[source_row]
            write_row({},[cell._style for cell in cells],auxiliary_dimensions[source_row])
            target_row=row_no-1
            for cell in cells:
                if cell.column in stale_columns:
                    continue
                target=ws.cell(target_row,cell.column)
                target.value=cell.value
                target.comment=copy(cell.comment)
                if cell.hyperlink:
                    target.hyperlink=copy(cell.hyperlink)
                    target.hyperlink.ref=target.coordinate
                if cell.data_type=="f":
                    # Moving an unknown formula can silently change its meaning.
                    # Preserve its cached reference value, or display the formula as text.
                    target.value=cached.get(cell.coordinate)
                    if target.value is None:
                        target.value=cell.value
                        target.data_type="s"
                    previous=target.comment.text+"\n" if target.comment else ""
                    target.comment=Comment(previous+f"전월 원본 참고 수식: {cell.value}\n자동 대사 계산에서 제외했습니다.","대사 안내")
            amount_cell=ws.cell(target_row,fields["amount"])
            if kind=="total":
                letter=amount_cell.column_letter
                end=header+body_rows
                if not body_rows:
                    amount_cell.value="=0"
                elif grouped:
                    label_col=ws.cell(header,fields[subtotal_label]).column_letter
                    amount_cell.value=f'=SUMIF({label_col}{header+1}:{label_col}{end},"소계",{letter}{header+1}:{letter}{end})'
                else:
                    amount_cell.value=f"=SUM({letter}{header+1}:{letter}{end})"
                amount_cell.comment=Comment("당월 명세서 전체 합계입니다. 거래처 소계를 중복 합산하지 않습니다.","대사 안내")
            elif kind=="held":
                if amount_cell.value is not None:
                    previous=amount_cell.comment.text+"\n" if amount_cell.comment else ""
                    amount_cell.comment=Comment(previous+f"이 행은 읽을 수 없어 대사에서 제외하고 그대로 두었습니다: {held[source_row]}","대사 안내")
                write_status(ws,target_row,status_column,fields["amount"],"확인 필요",
                             text=f"확인 필요 · 읽지 못해 그대로 둠\n{held[source_row]}")
            elif amount_cell.value is not None:
                previous=amount_cell.comment.text+"\n" if amount_cell.comment else ""
                amount_cell.comment=Comment(previous+"전월 원본의 참고 값입니다. 당월 더존 잔액으로 자동 확정한 값이 아니므로 재확인해 주세요.","대사 안내")
            row_text=" / ".join(str(cell.value) for cell in cells if cell.value is not None)
            if kind=="held":
                audits.append(("읽지 못한 행 보존","","",row_text,"",
                               f"{held[source_row]} · 대사 계산에서 제외하고 원본 서식으로 보존",
                               template_path.name,source_row))
                if not any(r[7]==source_row and r[6] in ("명세서",template_path.name) for r in reviews):
                    reviews.append(("입력 데이터 확인","","",row_text,"",held[source_row],"명세서",source_row))
            else:
                audits.append(("보조 행 보존","","",row_text,
                               "","대사 계산에서 제외하고 하단에 원본 서식으로 보존",template_path.name,source_row))
            source_targets[(template_path.name,source_row)]=[target_row]
        for region in auxiliary_merges:
            ws.merge_cells(start_row=auxiliary_targets[region.min_row],
                           end_row=auxiliary_targets[region.max_row],
                           start_column=region.min_col,end_column=region.max_col)
        if row_no != header+required_rows+1:
            raise AssertionError("당월 명세서 행 수와 실제 생성 행 수가 일치하지 않습니다.")

        # Original detail and payment rows remain traceable in this auxiliary tab.
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

        # Status is the sole optional extra column; detailed reasons stay removable.
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
        changes["J3"]=f"검토 필요 {len(reviews)}건 · RAW 전표 {plan.raw_count}건 · 일치 후보 {len(offsets)}건"
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

        # A new name avoids renaming a prior sheet, which would silently
        # re-point any formula that references it.
        balances=wb.create_sheet("당월 잔액검증")
        balances.append(["거래처코드","거래처명","전월 잔액","당월 대변(발생·조정)",
                         "당월 차변(지급·조정)","계산 잔액","생성 명세서 잔액","계산 차액",
                         "계산 검증","원장 잔액","원장 차액","원장 비교","검증 근거"])
        for record in plan.balances:
            balances.append(record)
        balances.freeze_panes="C2"
        for column in "ABCDEFGHIJKL":
            balances.column_dimensions[column].width=23
        balances.column_dimensions["M"].width=80
        for row in balances.iter_rows(min_row=2):
            for idx in (2,3,4,5,6,7,9,10):
                row[idx].number_format='#,##0;[Red](#,##0);"0"'
        balances.auto_filter.ref=balances.dimensions

        def link_source_rows(sheet):
            sheet["H1"].comment=Comment(
                "행 번호는 원본 파일 기준입니다. 클릭하면 생성된 당월 명세서의 해당 거래 행으로 이동합니다.",
                "대사 안내"
            )
            for row in sheet.iter_rows(min_row=2,max_col=8):
                source=row[6].value
                if source=="명세서":
                    source=template_path.name
                targets=source_targets.get((source,row[7].value),[])
                # Never guess when different inputs have the same file/row identity.
                if len(targets)!=1:
                    continue
                destination=f"{quote_sheetname(ws.title)}!{ws.cell(targets[0],fields['desc']).coordinate}"
                cell=row[7]
                cell.hyperlink=Hyperlink(ref=cell.coordinate,location=destination,
                                        tooltip="당월 명세서의 해당 거래로 이동")
                font=copy(cell.font)
                font.color="0563C1"
                font.underline="single"
                cell.font=font

        for extra in (offset_sheet,summary,changes):
            link_source_rows(extra)
        for generated in (offset_sheet,summary,changes,balances):
            force_text_formulas(generated)

        # Let Excel recalculate all copied per-vendor subtotal formulas on open.
        from openpyxl.workbook.properties import CalcProperties
        wb.calculation=CalcProperties(calcMode="auto",fullCalcOnLoad=True,forceFullCalc=True)
        wb.active=wb.worksheets.index(ws)
        wb.save(out)
        return notices
    finally:
        wb.close()
