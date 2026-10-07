from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

PALETTE={"header":"1D1D1F","carry":"FFF4CC","new":"E8F5EE","subtotal":"F2F2F7","total":"E5E5EA"}
AMOUNT_HEADERS={"금액","전월금액","더존차변","대변","차월 초안 금액"}
_THIN=Side(style="thin",color="E5E5EA")

def style_workbook(wb, draft, info):
    for sheet in wb.worksheets:
        sheet.sheet_view.showGridLines=False
        for cell in sheet[1]:
            cell.font=Font(bold=True,color="FFFFFF")
            cell.fill=PatternFill("solid",fgColor=PALETTE["header"])
            cell.alignment=Alignment(vertical="center",horizontal="center")
        sheet.freeze_panes="A2"
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment=Alignment(vertical="top",wrap_text=True)
                cell.border=Border(bottom=_THIN)
        for col_idx,cell in enumerate(sheet[1],1):
            if cell.value in AMOUNT_HEADERS:
                for rr in range(2,sheet.max_row+1):
                    sheet.cell(rr,col_idx).number_format="#,##0"
                    sheet.cell(rr,col_idx).alignment=Alignment(horizontal="right",vertical="top")
        for col in sheet.columns:
            sheet.column_dimensions[col[0].column_letter].width=min(max(max(len(str(x.value or "")) for x in col)+2,10),40)

    for rr in range(2,draft.max_row+1):
        kind=draft.cell(rr,1).value
        fill={"전월이월":PALETTE["carry"],"당월신규":PALETTE["new"],"소계":PALETTE["subtotal"],"전체합계":PALETTE["total"]}.get(kind)
        if fill:
            for cell in draft[rr]: cell.fill=PatternFill("solid",fgColor=fill)
        if kind in ("소계","전체합계"):
            for cell in draft[rr]: cell.font=Font(bold=True)
    draft.column_dimensions["E"].width=36
    draft.column_dimensions["G"].width=22
    draft.column_dimensions["H"].width=28
    info.column_dimensions["A"].width=22
    info.column_dimensions["B"].width=58
