"""Optional status column on the generated statement; never edit Douzone RAW.

The source statement's business columns are untouched. Status values are
generated to the right of all occupied source columns and can be deleted by
the person reviewing the draft.
"""
from copy import copy

from openpyxl.styles import Alignment, PatternFill
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter


_STATUS_COLORS={
    "자동 대사":"E2F0D9",
    "전월 이월":"E9EDF4",
    "순잔액 대사":"E2F0D9",
    "당월 신규":"EEE3FF",
    "당월 지급":"E6F2FF",
    "원장 단독":"FFF2CC",
    "확인 필요":"FFE4CC",
    "오타 의심":"FADBD8",
}
_DISPLAY={
    "대사 일치":"자동 대사",
    "당월 발생":"당월 신규",
}


def find_status_column(sheet, header_row):
    """Use existing status column, or append without overwriting owner columns.

    A new column is F when E is the last occupied business column, but is
    farther right when the source includes notes or other user data.
    """
    existing=[
        cell.column for cell in sheet[header_row]
        if str(cell.value or "").strip().replace(" ","")=="처리상태"
    ]
    if len(existing)>1:
        raise ValueError("처리상태 열이 둘 이상 있어 어느 열을 사용할지 판단할 수 없습니다.")
    if existing:
        return existing[0]

    last=max(
        (cell.column for cell in sheet._cells.values()
         if cell.row>=header_row and cell.value is not None),
        default=0,
    )
    if not last:
        raise ValueError("명세서의 데이터 열을 확인할 수 없습니다.")
    return last+1


def prepare_status_header(sheet, header_row, status_col, example_header_col):
    target=sheet.cell(header_row,status_col)
    if target.value not in (None,"처리상태"):
        raise ValueError("처리상태 표시 위치에 기존 데이터가 있어 자동 작성할 수 없습니다.")
    if target.value is None:
        target._style=copy(sheet.cell(header_row,example_header_col)._style)
    target.value="처리상태"
    letter=get_column_letter(status_col)
    if sheet.column_dimensions[letter].width==13.0:
        sheet.column_dimensions[letter].width=32


def write_status(sheet, row, column, amount_column, status, *, subtotal=False, reason=""):
    cell=sheet.cell(row,column)
    # Row styles are cloned by the Excel writer. The status cell receives the
    # same base font/border, before only its status-specific fill is applied.
    # Always clone the source detail/subtotal style; a pre-existing F-cell
    # format must not introduce mismatched fonts or borders in generated rows.
    cell._style=copy(sheet.cell(row,amount_column)._style)
    if subtotal:
        cell.value="소계"
        return

    label=_DISPLAY.get(status,status or "확인 필요")
    cell.value=label
    if reason:
        topics=[]
        for words,title in (
            (("거래처코드 차이","코드 오타","다른 거래처코드"),"거래처코드 확인"),
            (("거래처명 오타","거래처명 차이","원장 거래처명"),"거래처명 차이"),
            (("적요",),"적요 확인"),
            (("부분지급",),"부분지급"),
            (("수식",),"금액 수식 확인"),
            (("입력 데이터 확인","입력 확인"),"입력 오류로 자동 확정 보류"),
            (("중복","분할","합산"),"중복·분할 확인"),
            (("음수",),"음수 잔액 확인"),
            (("개별 전월 발생건과 연결 미확정",),"전월 발생건 연결 미확정"),
            (("대응 건 없음",),"전월 대응 건 없음"),
        ):
            if any(word in reason for word in words):
                topics.append(title)
        cell.value=label+" · "+(", ".join(topics) if topics else reason)
        cell.comment=Comment(reason,"대사 검토")
    cell.number_format="General"
    cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    cell.fill=PatternFill(fill_type="solid",fgColor=_STATUS_COLORS.get(label,"FFFFFF"))
