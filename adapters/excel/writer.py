from pathlib import Path
from typing import Iterable
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from domain.models import Status


def write_result(path, results, new_items, issues, source_paths:Iterable, period_label:str):
    out=Path(path).resolve()
    if any(out==Path(p).resolve() for p in source_paths):
        raise ValueError("원본 Excel에는 저장할 수 없습니다.")
    wb=Workbook(); ws=wb.active; ws.title="확인필요"
    ws.append(["상태","사유","담당자","원본파일","원본시트","원본행","거래처코드","거래처명","전월적요","전월금액","더존거래처","더존적요","더존차변","더존행"])
    for x in results:
        if x.status==Status.MATCHED: continue
        j=x.journal
        ws.append([x.status.value,x.reason,x.prior.source.owner,x.prior.source.file_name,x.prior.source.sheet,x.prior.source.row,x.prior.vendor_code,x.prior.vendor_name,x.prior.description,int(x.prior.amount), j.vendor_name if j else "",j.description if j else "",int(j.debit) if j else "",j.row_number if j else ""])
    iq=wb.create_sheet("입력데이터확인"); iq.append(["출처","시트","행","필드","원본값","사유"])
    for x in issues: iq.append([x.source,x.sheet,x.row,x.field,str(x.raw_value),x.reason])
    # Draft is intentionally conservative: only definitely-unpaid prior items and
    # current-month credits are included. Review/ambiguous items stay out until a human decides.
    draft=wb.create_sheet("차월명세서 초안")
    draft.append(["구분","거래처코드","거래처명","날짜","적요","금액","근거","원본"])
    draft_rows=[]
    for x in results:
        if x.status==Status.UNPAID:
            p=x.prior
            draft_rows.append(["전월이월",p.vendor_code,p.vendor_name,p.date,p.description,int(p.amount),
                               "당월 대응 차변 없음",f"{p.source.file_name} · {p.source.row}행"])
    for j in new_items:
        draft_rows.append(["당월신규",j.vendor_code,j.vendor_name,j.date,j.description,int(j.credit),
                           "당월 미지급금 대변",f"더존 · {j.row_number}행"])
    draft_rows.sort(key=lambda x:(x[2] or "",x[1] or "",x[3] or "",x[4] or ""))

    current=None; subtotal=0
    for row in draft_rows:
        vendor=(row[1],row[2])
        if current is not None and vendor != current:
            draft.append(["소계",current[0],current[1],"","",subtotal,"거래처 소계",""])
            subtotal=0
        draft.append(row); subtotal += row[5]; current=vendor
    if current is not None:
        draft.append(["소계",current[0],current[1],"","",subtotal,"거래처 소계",""])
    if draft_rows:
        draft.append(["전체합계","","","","",sum(r[5] for r in draft_rows),"차월명세서 초안 합계",""])


    nw=wb.create_sheet("신규미지급"); nw.append(["기표일자","계정코드","거래처코드","거래처명","적요","대변"])
    for j in new_items: nw.append([j.date,j.account_code,j.vendor_code,j.vendor_name,j.description,int(j.credit)])
    ok=wb.create_sheet("자동대사완료"); ok.append(["상태","거래처코드","거래처명","전월적요","금액","더존행"])
    for x in results:
        if x.status==Status.MATCHED: ok.append([x.status.value,x.prior.vendor_code,x.prior.vendor_name,x.prior.description,int(x.prior.amount),x.journal.row_number])
    info=wb.create_sheet("요약",0)
    info.append(["대상 회계월",period_label])
    info.append(["확인 필요",sum(x.status!=Status.MATCHED for x in results)])
    info.append(["입력 형식 확인",len(issues)])
    info.append(["차월 초안",len(draft_rows)])
    info.append(["차월 초안 금액",sum(r[5] for r in draft_rows)])
    info.append(["초안 제외 검토건",sum(x.status not in (Status.MATCHED,Status.UNPAID) for x in results)])
    info.append(["초안 원칙","확정 미지급 이월 + 당월 신규만 포함 / 검토 필요 건은 제외"])

    palette={"header":"1D1D1F","carry":"FFF4CC","new":"E8F5EE","subtotal":"F2F2F7","total":"E5E5EA"}
    thin=Side(style="thin",color="E5E5EA")
    for sheet in wb.worksheets:
        sheet.sheet_view.showGridLines=False
        for cell in sheet[1]:
            cell.font=Font(bold=True,color="FFFFFF")
            cell.fill=PatternFill("solid",fgColor=palette["header"])
            cell.alignment=Alignment(vertical="center",horizontal="center")
        sheet.freeze_panes="A2"
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment=Alignment(vertical="top",wrap_text=True)
                cell.border=Border(bottom=thin)
        for col_idx,cell in enumerate(sheet[1],1):
            if cell.value in {"금액","전월금액","더존차변","대변","차월 초안 금액"}:
                for rr in range(2,sheet.max_row+1):
                    sheet.cell(rr,col_idx).number_format="#,##0"
                    sheet.cell(rr,col_idx).alignment=Alignment(horizontal="right",vertical="top")
        for col in sheet.columns:
            sheet.column_dimensions[col[0].column_letter].width=min(max(max(len(str(x.value or "")) for x in col)+2,10),40)

    for rr in range(2,draft.max_row+1):
        kind=draft.cell(rr,1).value
        fill={"전월이월":palette["carry"],"당월신규":palette["new"],"소계":palette["subtotal"],"전체합계":palette["total"]}.get(kind)
        if fill:
            for cell in draft[rr]: cell.fill=PatternFill("solid",fgColor=fill)
        if kind in ("소계","전체합계"):
            for cell in draft[rr]: cell.font=Font(bold=True)
    draft.column_dimensions["E"].width=36
    draft.column_dimensions["G"].width=22
    draft.column_dimensions["H"].width=28
    info.column_dimensions["A"].width=22
    info.column_dimensions["B"].width=58
    wb.save(out)
