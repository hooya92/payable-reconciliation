import math
import unicodedata

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

PALETTE={"header":"1D1D1F","carry":"FFF4CC","new":"E8F5EE","subtotal":"F2F2F7","total":"E5E5EA"}
AMOUNT_HEADERS={"금액","명세서금액","더존차변","대변","차월 초안 금액","차월 이월 초안 금액"}
_THIN=Side(style="thin",color="E5E5EA")


def _display_width(value):
    if value is None:
        return 0
    return sum(
        2 if unicodedata.east_asian_width(ch) in ("W","F") else 1
        for ch in str(value)
    )


def _autofit_sheet(sheet):
    widths={}
    for col in sheet.columns:
        letter=col[0].column_letter
        header=str(col[0].value or "")
        content_width=max((_display_width(cell.value) for cell in col),default=0)
        if any(key in header for key in ("적요","사유","참고","원본","거래처명")):
            min_width,max_width=24,48
        elif any(key in header for key in ("코드","행","상태","구분")):
            min_width,max_width=10,22
        else:
            min_width,max_width=10,34
        width=min(max(content_width+3,min_width),max_width)
        sheet.column_dimensions[letter].width=width
        widths[letter]=width

    sheet.row_dimensions[1].height=max(sheet.row_dimensions[1].height or 0,24)
    for row_idx in range(2,sheet.max_row+1):
        lines=1
        for cell in sheet[row_idx]:
            if cell.value in (None,""):
                continue
            usable=max(widths.get(cell.column_letter,10)-2,1)
            needed=sum(
                max(1,math.ceil(_display_width(part)/usable))
                for part in str(cell.value).splitlines()
            )
            lines=max(lines,needed)
        sheet.row_dimensions[row_idx].height=min(max(20,lines*18),90)


def style_workbook(wb, draft, info):
    for sheet in wb.worksheets:
        sheet.sheet_view.showGridLines=False
        for cell in sheet[1]:
            cell.font=Font(bold=True,color="FFFFFF")
            cell.fill=PatternFill("solid",fgColor=PALETTE["header"])
            cell.alignment=Alignment(vertical="center",horizontal="center")
        sheet.freeze_panes="A2"
        sheet.auto_filter.ref=sheet.dimensions
        sheet.sheet_properties.pageSetUpPr.fitToPage=True
        sheet.page_setup.fitToWidth=1
        sheet.page_setup.fitToHeight=0
        sheet.page_margins.left=0.25
        sheet.page_margins.right=0.25
        sheet.page_margins.top=0.5
        sheet.page_margins.bottom=0.5
        sheet.oddFooter.center.text="페이지 &P / &N"
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment=Alignment(vertical="top",wrap_text=True)
                cell.border=Border(bottom=_THIN)
        for col_idx,cell in enumerate(sheet[1],1):
            if cell.value in AMOUNT_HEADERS:
                for rr in range(2,sheet.max_row+1):
                    sheet.cell(rr,col_idx).number_format="#,##0"
                    sheet.cell(rr,col_idx).alignment=Alignment(horizontal="right",vertical="top",wrap_text=True)
        _autofit_sheet(sheet)

    for rr in range(2,draft.max_row+1):
        kind=draft.cell(rr,1).value
        fill={"전월이월":PALETTE["carry"],"미지급이월":PALETTE["carry"],"당월신규":PALETTE["new"],"소계":PALETTE["subtotal"],"전체합계":PALETTE["total"]}.get(kind)
        if fill:
            for cell in draft[rr]: cell.fill=PatternFill("solid",fgColor=fill)
        if kind in ("소계","전체합계"):
            for cell in draft[rr]: cell.font=Font(bold=True)
    draft.sheet_view.showGridLines=False
    draft.page_setup.orientation="landscape"
    draft.print_title_rows="1:1"
    draft.column_dimensions["E"].width=max(draft.column_dimensions["E"].width or 0,36)
    draft.column_dimensions["G"].width=max(draft.column_dimensions["G"].width or 0,22)
    draft.column_dimensions["H"].width=max(draft.column_dimensions["H"].width or 0,28)
    info.freeze_panes=None
    info.auto_filter.ref=None
    info.row_dimensions[1].height=24
    for rr in range(1,info.max_row+1):
        info.cell(rr,1).font=Font(bold=True)
        info.cell(rr,2).alignment=Alignment(vertical="center",wrap_text=True)
    info.column_dimensions["A"].width=22
    info.column_dimensions["B"].width=58
