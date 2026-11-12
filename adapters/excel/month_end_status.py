"""Optional status column on the generated statement; never edit Douzone RAW.

The source statement's business columns are untouched. Status values are
generated to the right of all occupied source columns and can be deleted by
the person reviewing the draft.
"""
from copy import copy
from decimal import Decimal

from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
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
    "오발행·취소 의심":"FFE4CC",
}
_DISPLAY={
    "대사 일치":"자동 대사",
    "당월 발생":"당월 신규",
}


# Words this program itself writes into the status column. Used to recognise a
# status column left over from the previous month's output, which must not be
# carried into the new month as if it were the owner's own notes.
_GENERATED_LABELS=(
    "자동 대사","순잔액 대사","전월 이월","당월 신규","당월 지급","원장 단독",
    "확인 필요","오타 의심","오발행·취소 의심","대사 일치","당월 발생","소계",
)


def is_generated_status(value):
    """True for text this program wrote as a row status (never for free-form notes)."""
    if not isinstance(value,str):
        return False
    text=value.strip()
    return any(
        text==label or text.startswith((label+" ·",label+" ↔",label+" 0"))
        for label in _GENERATED_LABELS
    )


_GREEN="1B7233"
_AMBER="8A5A00"
_RED="B42318"
_GRAY="5F6368"
_DARK="1D1D1F"


def _num(value):
    value=Decimal(value)
    text=f"{abs(value):,.0f}" if value==value.to_integral_value() else f"{abs(value):,.2f}"
    return ("−" if value<0 else "")+text


def subtotal_note_text(note, font_name="맑은 고딕", font_size=10):
    """Vendor subtotal explanation as one rich-text cell (bold / colour / line breaks).

    The whole explanation lives in this single status cell so that deleting the
    status column removes all of it and nothing is added around the statement.
    """
    def run(text, color=_DARK, bold=False):
        return TextBlock(InlineFont(rFont=font_name,sz=font_size,b=bold or None,color=color),text)

    terms=[("전월 합계",note.opening,"")]
    if note.credit:
        terms.append(("+ 당월 대변",note.credit,""))
    if note.debit:
        terms.append(("− 당월 차변",note.debit,""))

    if note.state=="mismatch":
        head=("✗ 검산 불일치",_RED)
    elif note.state=="hold":
        head=("△ 검산 보류 · 입력 확인 필요",_AMBER)
    elif note.review:
        head=("△ 검산 일치 · 확인 필요 항목 있음",_AMBER)
    else:
        head=("✓ 검산 일치",_GREEN)

    if note.closing==0:
        # Documented rule: an exactly-zero subtotal is always marked "소계 0".
        head=("소계 0 · "+head[0],head[1])
    parts=[]
    if len(terms)<3 and not note.groups and note.state=="ok" and not note.review:
        # Short vendor: the whole explanation on one line.
        pieces=[f"전월 {_num(note.opening)}"]
        if note.credit:
            pieces.append(f"+ 대변 {_num(note.credit)}")
        if note.debit:
            pieces.append(f"− 차변 {_num(note.debit)}")
        return CellRichText(
            run(head[0],head[1],True),
            run("  "+" ".join(pieces)+f" = {_num(note.closing)}",_GRAY),
        )
    parts.append(run(head[0]+"\n",head[1],True))
    for label,value,_ in terms:
        parts+=[run(f"{label}  ",_GRAY),run(_num(value)+"\n")]
    parts+=[run("= 말 잔액  ",_GRAY),run(_num(note.closing),bold=True)]
    if note.state=="mismatch":
        parts+=[run(f"\n계산상 잔액  {_num(note.expected)}",_RED)]
    for group in note.groups:
        remaining=group.remaining
        if remaining==0:
            tail=("✓ 일치 (후보)",_GREEN,"=")
        elif remaining>0:
            tail=(f"△ 남음 {_num(remaining)} (후보)",_AMBER,"−")
        else:
            tail=(f"△ 초과 {_num(-remaining)} (후보)",_AMBER,"−")
        parts+=[run("\n\n"),run(f"{group.month}월 묶음  ",bold=True),run(tail[0]+"\n",tail[1],True),
                run(f"청구 {_num(group.claims)} {tail[2]} 지급 {_num(group.payments)}",_GRAY)]
    return CellRichText(*parts)


def note_line_count(rich):
    return str(rich).count("\n")+1


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


def write_status(sheet, row, column, amount_column, status, *, subtotal=False, reason="", subtotal_zero=False,
                 text=None, note=None):
    cell=sheet.cell(row,column)
    # Row styles are cloned by the Excel writer. The status cell receives the
    # same base font/border, before only its status-specific fill is applied.
    # Always clone the source detail/subtotal style; a pre-existing F-cell
    # format must not introduce mismatched fonts or borders in generated rows.
    cell._style=copy(sheet.cell(row,amount_column)._style)
    if subtotal:
        cell.value="소계 0" if subtotal_zero else "소계"
        if subtotal_zero:
            amount=sheet.cell(row,amount_column)
            # Replace only the zero section, keeping positive/negative formatting.
            sections=[]
            start=0
            quoted=escaped=False
            for idx,char in enumerate(amount.number_format):
                if escaped:
                    escaped=False
                elif char in ("\\","_","*"):
                    escaped=True
                elif char=='"':
                    quoted=not quoted
                elif char==";" and not quoted:
                    sections.append(amount.number_format[start:idx])
                    start=idx+1
            sections.append(amount.number_format[start:])
            if len(sections)>=3:
                sections[2]='"0"'
                amount.number_format=";".join(sections)
        if note is not None:
            rich=subtotal_note_text(note,cell.font.name or "맑은 고딕",cell.font.sz or 10)
            cell.value=rich
            cell.alignment=Alignment(horizontal="left",vertical="top",wrap_text=True)
            # Excel fits a row without a fixed height by itself; only a template
            # row that fixes its height smaller than the text needs is enlarged.
            needed=15*note_line_count(rich)+3
            current=sheet.row_dimensions[row].height
            if current is not None and current<needed:
                sheet.row_dimensions[row].height=needed
        return

    label=_DISPLAY.get(status,status or "확인 필요")
    cell.value=label
    if reason:
        topics=[]
        for words,title in (
            (("오발행","취소 의심","역분개","수정 표시"),"취소·재발행 확인"),
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
    if text:
        # Explicit wording (what the row was offset against); the generated
        # reason above stays in the cell comment.
        cell.value=text
    cell.number_format="General"
    cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    cell.fill=PatternFill(fill_type="solid",fgColor=_STATUS_COLORS.get(label,"FFFFFF"))
