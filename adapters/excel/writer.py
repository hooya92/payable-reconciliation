from pathlib import Path
from typing import Iterable
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
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
    nw=wb.create_sheet("신규미지급"); nw.append(["기표일자","계정코드","거래처코드","거래처명","적요","대변"])
    for j in new_items: nw.append([j.date,j.account_code,j.vendor_code,j.vendor_name,j.description,int(j.credit)])
    ok=wb.create_sheet("자동대사완료"); ok.append(["상태","거래처코드","거래처명","전월적요","금액","더존행"])
    for x in results:
        if x.status==Status.MATCHED: ok.append([x.status.value,x.prior.vendor_code,x.prior.vendor_name,x.prior.description,int(x.prior.amount),x.journal.row_number])
    info=wb.create_sheet("요약",0); info.append(["대상 회계월",period_label]); info.append(["확인 필요",sum(x.status!=Status.MATCHED for x in results)]); info.append(["입력 형식 확인",len(issues)])
    for sheet in wb.worksheets:
        if sheet.max_row and sheet.max_column:
            for c in sheet[1]:
                c.font=Font(bold=True); c.fill=PatternFill("solid",fgColor="E9EEF5"); c.alignment=Alignment(vertical="center")
            sheet.freeze_panes="A2"
            for col in sheet.columns:
                sheet.column_dimensions[col[0].column_letter].width=min(max(max(len(str(c.value or "")) for c in col)+2,10),44)
    wb.save(out)
