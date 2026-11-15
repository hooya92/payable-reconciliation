"""The two companion sheets next to the generated month-end statement: 검토필요 and 대사상세.

The statement sheet itself is never styled here. These sheets are removable review aids:
a title line, a one-line summary, then a styled table (header on row 4, data from row 5).
"""
from __future__ import annotations

from collections import Counter

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, quote_sheetname

REVIEW_SHEET = "검토필요"
DETAIL_SHEET = "대사상세"
REVIEW_MARKER = "MONTH_END_DRAFT_REVIEW"      # J1 of 검토필요 (read by the next month's guard)
DETAIL_MARKER = "MONTH_END_DRAFT_DETAIL"      # L1 of 대사상세
HEADER_ROW = 4

REVIEW_HEADERS = ["검토 상태", "거래처코드", "거래처명", "적요·원본값", "금액", "사유", "원본파일", "원본행"]
DETAIL_HEADERS = ["구분", "거래처코드", "거래처명", "명세 내용", "금액", "근거", "원본파일", "원본행", "대사 상대", "본문"]

# Header rows of companion sheets written by earlier versions (row 1, first cells).
_LEGACY_HEADERS = (
    ("구분", "거래처코드", "거래처명", "원본 적요", "발생 금액"),            # 대사내역
    ("구분", "거래처코드", "거래처명", "명세 내용", "금액", "근거"),        # 변경내역
    ("검토 상태", "거래처코드", "거래처명", "적요·원본값", "금액", "사유"),  # 검토필요
    ("거래처코드", "거래처명", "전월 잔액", "당월 대변(발생·조정)"),         # 당월 잔액검증
)

_FONT = "맑은 고딕"
_NAVY = "1F3A5F"
_LINE = Side(style="thin", color="D9DCE1")
_BORDER = Border(top=_LINE, bottom=_LINE, left=_LINE, right=_LINE)
_HEAD_BORDER = Border(left=Side(style="thin", color="2E4D73"), right=Side(style="thin", color="2E4D73"))
_ZEBRA = (PatternFill("solid", fgColor="FFFFFF"), PatternFill("solid", fgColor="F7F8FA"))
_TEXT = Font(name=_FONT, size=10, color="1D1D1F")
_MONEY = '#,##0;[Red]-#,##0;"-"'

# status word -> (background, text colour)
_CHIPS = {
    "확인 필요": ("FCE4D6", "B5491B"), "검토 보류": ("FCE4D6", "B5491B"), "잔액 확인": ("FCE4D6", "B5491B"),
    "입력 데이터 확인": ("FCE4D6", "B5491B"), "원장 단독": ("FFF2CC", "8A6D00"),
    "대사 일치": ("E2EFDA", "2E6B30"), "전월 이월": ("EDEDED", "555555"), "당월 발생": ("DDEBF7", "1F4E79"),
    "보조 행 보존": ("F2F2F2", "777777"), "읽지 못한 행 보존": ("FFF2CC", "8A6D00"),
}
_CHIP_DEFAULT = ("EEF0F3", "333333")


def is_generated_companion(ws) -> bool:
    """True for a review/detail sheet this program wrote earlier (any version)."""
    if ws.cell(1, 10).value == REVIEW_MARKER or ws.cell(1, 12).value == DETAIL_MARKER:
        return True
    first = tuple(ws.cell(1, col).value for col in range(1, 7))
    return any(first[:len(sig)] == sig for sig in _LEGACY_HEADERS)


def review_data_rows(ws) -> int:
    """Number of review entries left in a generated 검토필요 sheet (old or new layout)."""
    header = next((r for r in range(1, min(ws.max_row, 10) + 1)
                   if ws.cell(r, 1).value == REVIEW_HEADERS[0]), 1)
    count = 0
    for row in ws.iter_rows(min_row=header + 1, max_col=8, values_only=True):
        if any(v not in (None, "") for v in row):
            count += 1
    return count


def _display(text):
    return text.replace(" · ", "\n") if isinstance(text, str) else text


def _layout(ws, title, subtitle, headers, rows, widths, *, money, chip, wrap, tab, ok=False):
    ws.sheet_properties.tabColor = tab
    ws.sheet_view.showGridLines = False
    ws["A1"] = title
    ws["A1"].font = Font(name=_FONT, size=15, bold=True, color="2E6B30" if ok else "1D1D1F")
    ws["A2"] = subtitle
    ws["A2"].font = Font(name=_FONT, size=10, color="2E6B30" if ok else "6E6E73")
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 6
    head_font = Font(name=_FONT, size=10, bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor=_NAVY)
    for col, name in enumerate(headers, 1):
        cell = ws.cell(HEADER_ROW, col, name)
        cell.font, cell.fill, cell.border = head_font, head_fill, _HEAD_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[HEADER_ROW].height = 30
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width

    centred = {col for col, name in enumerate(headers, 1)
               if col == chip or any(k in name for k in ("코드", "행", "일자"))}
    aligns = {}
    for col in range(1, len(headers) + 1):
        if col in money:
            aligns[col] = Alignment(horizontal="right", vertical="center", indent=1)
        elif col in centred:
            aligns[col] = Alignment(horizontal="center", vertical="center", wrap_text=col in wrap)
        else:
            aligns[col] = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=col in wrap)
    chip_styles = {}
    for offset, values in enumerate(rows):
        r = HEADER_ROW + 1 + offset
        fill = _ZEBRA[offset % 2]
        lines = 1
        for col, value in enumerate(values[:len(headers)], 1):
            if col == 6:   # 사유 / 근거: one clause per line
                value = _display(value)
            cell = ws.cell(r, col, value)
            cell.font, cell.fill, cell.border, cell.alignment = _TEXT, fill, _BORDER, aligns[col]
            if col in money and isinstance(value, (int, float)):
                cell.number_format = _MONEY
            if col == chip and value:
                if value not in chip_styles:
                    bg, fg = _CHIPS.get(value, _CHIP_DEFAULT)
                    chip_styles[value] = (PatternFill("solid", fgColor=bg),
                                          Font(name=_FONT, size=10, bold=True, color=fg))
                cell.fill, cell.font = chip_styles[value]
            if col in wrap and isinstance(value, str):
                per_line = max(int(widths[col - 1] * 0.9), 8)
                lines = max(lines, sum(len(part) // per_line + 1 for part in value.split("\n")))
        if lines > 1:
            ws.row_dimensions[r].height = 15 * lines + 5
    ws.freeze_panes = f"A{HEADER_ROW + 1}"
    if rows:
        ws.auto_filter.ref = f"A{HEADER_ROW}:{get_column_letter(len(headers))}{HEADER_ROW + len(rows)}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{HEADER_ROW}:{HEADER_ROW}"


def write_review_sheet(wb, reviews, statement_title):
    ws = wb.create_sheet(REVIEW_SHEET)
    n = len(reviews)
    if n:
        title = f"검토필요  ·  {n}건"
        subtitle = (f"확인 후 '{statement_title}' 본문을 고치고, 이 시트의 항목을 모두 지우거나 시트를 삭제해야 "
                    "다음 달 대사에 쓸 수 있습니다.  원본행을 누르면 본문 해당 행으로 이동합니다.")
    else:
        title = "검토필요  ·  0건  ✓"
        subtitle = "검토할 항목이 없습니다. 이 시트는 지워도 됩니다."
    _layout(ws, title, subtitle, REVIEW_HEADERS, [list(r) for r in reviews],
            [14, 12, 14, 28, 15, 62, 14, 8], money={5}, chip=1, wrap={4, 6}, tab="E26B1E", ok=not n)
    ws.cell(1, 10, REVIEW_MARKER)
    ws.column_dimensions["J"].hidden = True
    return ws


def _counterparts(offsets):
    """Text naming the other side of each matched pair, keyed by the row it belongs to."""
    by_prior, by_raw = {}, {}
    for o in offsets:
        prior_key = (o[6], o[7])
        by_prior[prior_key] = f"↔ 더존 {o[12]}행 · {o[9] or ''} · {o[8] or ''}".replace(" ·  · ", " · ").rstrip(" ·")
        raw_key = (o[12], -o[10] if isinstance(o[10], (int, float)) else None)
        by_raw.setdefault(raw_key, []).append(f"↔ 명세서 {o[7]}행 · {o[5] or ''} · {o[3] or ''}".rstrip(" ·"))
    return by_prior, by_raw


def write_detail_sheet(wb, audits, offsets, review_count, raw_count, prior_name):
    ws = wb.create_sheet(DETAIL_SHEET)
    by_prior, by_raw = _counterparts(offsets)
    rows = []
    for record in audits:
        record = list(record)
        pair = ""
        if record[0] == "대사 일치":
            if record[6] == prior_name:
                pair = by_prior.get((record[6], record[7]), "")
            else:
                found = by_raw.get((record[7], record[4]), [])
                pair = found[0] if len(found) == 1 else ""
        rows.append(record + [pair, None])
    counts = Counter(r[0] for r in rows)
    summary = "  ·  ".join(f"{k} {v}" for k, v in counts.items())
    subtitle = (f"{summary}   |   검토 필요 {review_count}건 · RAW 전표 {raw_count}건 · 대사 짝 {len(offsets)}건.  "
                "발생·지급은 모두 보존하며, 대사 일치는 실제 상계 확정과 다릅니다.")
    _layout(ws, f"대사상세  ·  {len(rows)}행", subtitle, DETAIL_HEADERS, rows,
            [14, 12, 14, 26, 15, 40, 14, 8, 36, 20], money={5}, chip=1, wrap={4, 6, 9}, tab="9AA5B1")
    ws.cell(1, 12, DETAIL_MARKER)
    ws.column_dimensions["L"].hidden = True
    return ws


def table_rows(ws):
    """The data rows of a companion sheet (below its header row), as value tuples."""
    return [row[:len(DETAIL_HEADERS)] for row in ws.iter_rows(min_row=HEADER_ROW + 1, values_only=True)]


KEPT, DONE, ONE_SIDE, WRONG = "본문에 있음", "✓ 짝과 함께 지움", "⚠ 한쪽만 지움", "⚠ 지움 (이월·신규 행)"
DONE_SINGLE, REVIEWED = "✓ 지움", "지움 (검토 건)"
_MUST_STAY = ("전월 이월", "당월 발생")


def add_presence_column(ws, statement, locate, desc_col, amount_col, offsets, prior_name):
    """Column J of 대사상세 follows the statement live: when a person deletes (or clears) a row there,
    the formula's reference breaks and the cell says so. A matched pair should be removed together;
    a carried-forward or new row should stay. Written after force_text_formulas: these are ours."""
    sheet = quote_sheetname(statement.title)

    def gone(target):
        d = f"{sheet}!${get_column_letter(desc_col)}${target}"
        a = f"{sheet}!${get_column_letter(amount_col)}${target}"
        # COUNTA never errors on a cell's own error value; it errors only on a deleted reference.
        return f"IFERROR(COUNTA({d},{a})=0,TRUE)"

    rows = list(ws.iter_rows(min_row=HEADER_ROW + 1, max_col=len(DETAIL_HEADERS)))
    by_key = {}
    for row in rows:
        by_key.setdefault((row[6].value, row[7].value, row[4].value), []).append(row)
    partner_of = {}
    for o in offsets:
        if not isinstance(o[10], (int, float)):
            continue
        prior_side = by_key.get((o[6], o[7], o[4]), [])
        raw_side = [r for (f, n, amt), found in by_key.items() if f != prior_name and n == o[12]
                    and amt == -o[10] for r in found]
        if len(prior_side) == 1 and len(raw_side) == 1:
            p, r = prior_side[0], raw_side[0]
            partner_of[id(p)] = (r[6].value, r[7].value)
            partner_of[id(r)] = (p[6].value, p[7].value)

    for row in rows:
        cell = row[9]
        target = locate(row[6].value, row[7].value)
        if target is None or row[0].value == "보조 행 보존":
            continue
        mine = gone(target)
        if row[0].value == "대사 일치":
            other = partner_of.get(id(row))
            other_target = locate(*other) if other else None
            if other_target is not None:
                theirs = gone(other_target)
                cell.value = (f'=IF(AND({mine},{theirs}),"{DONE}",IF(OR({mine},{theirs}),"{ONE_SIDE}","{KEPT}"))')
            else:
                cell.value = f'=IF({mine},"{DONE_SINGLE}","{KEPT}")'
        else:
            removed = WRONG if row[0].value in _MUST_STAY else REVIEWED
            cell.value = f'=IF({mine},"{removed}","{KEPT}")'
        cell.alignment = Alignment(horizontal="center", vertical="center")

    last = HEADER_ROW + max(len(rows), 1)
    col = f"$J${HEADER_ROW + 1}:$J${last}"
    kind = f"$A${HEADER_ROW + 1}:$A${last}"
    ws.conditional_formatting.add(f"J{HEADER_ROW + 1}:J{last}", FormulaRule(
        formula=[f'LEFT(J{HEADER_ROW + 1},1)="✓"'], stopIfTrue=True,
        fill=PatternFill("solid", fgColor="E2EFDA"), font=Font(name=_FONT, bold=True, color="2E6B30")))
    ws.conditional_formatting.add(f"J{HEADER_ROW + 1}:J{last}", FormulaRule(
        formula=[f'LEFT(J{HEADER_ROW + 1},1)="⚠"'], stopIfTrue=True,
        fill=PatternFill("solid", fgColor="FCE4D6"), font=Font(name=_FONT, bold=True, color="B5491B")))
    # Live summary under the title: what is left to tidy up in the statement.
    static = str(ws["A2"].value or "").replace('"', '""')
    ws["A2"] = (f'="본문 정리  ·  완료 "&COUNTIF({col},"✓*")&"행 · 남은 대사 짝 "'
                f'&COUNTIFS({kind},"대사 일치",{col},"{KEPT}")&"행 · 주의 "&COUNTIF({col},"⚠*")&"행'
                f'   |   {static}"')
