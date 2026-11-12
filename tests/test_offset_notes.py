"""Status-column wording that says what each row was offset against (display only)."""
import tempfile
import unittest
from decimal import Decimal as D
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

from adapters.excel.month_end_statement import write_month_end_statement
from adapters.excel.month_end_status import is_generated_status, subtotal_note_text
from application.service import run_reconciliation
from domain.offset_hints import MonthGroup, VendorNote, month_groups, month_token
from tests import test_journal_corrections as fixtures

PERIOD=fixtures.PERIOD


def raw_file(path, rows):
    wb=Workbook(); ws=wb.active
    ws.append(["기표일자","계정코드","계정과목명","거래처코드","거래처명","적요","차변","대변"])
    for code,name,desc,debit,credit,when in rows:
        ws.append([when,"25301","미지급금",code,name,desc,debit,credit])
    wb.save(path); wb.close()


def net_vendor_prior(path, stale=False):
    """One vendor with 7월/8월 claims and a negative 7월 adjustment (a signed-net vendor)."""
    wb=Workbook(); ws=wb.active; ws.title="26.08"
    ws.append(["미지급금 명세"]); ws.append(["코드","거래처","날짜","적요","금액"])
    ws.append([None,None,"2026-08-01","7월 위탁물류비 A",100])
    ws.append([None,None,"2026-08-01","7월 위탁물류비 B",200])
    ws.append([None,None,"2026-08-01","계정정리)7월 패널티공제",-10])
    ws.append([None,None,"2026-08-31","8월 위탁물류비 C",300])
    ws.append(["001111","가상물류","소계",None,"=SUM(E3:E6)"])
    ws.append(["합 계",None,None,None,'=SUMIF(C3:C7,"소계",E3:E7)'])
    for col in range(1,6):
        ws.cell(7,col).fill=PatternFill("solid",fgColor="FFFF00")
    if stale:
        # last month's own status words left in two outside columns, with their fills
        for row in (3,4):
            ws.cell(row,9,"자동 대사"); ws.cell(row,9).fill=PatternFill("solid",fgColor="E2F0D9")
        ws.cell(7,10,"소계"); ws.cell(7,10).fill=PatternFill("solid",fgColor="FFFF00")
        ws.cell(2,11,"비고"); ws.cell(3,11,"담당자 메모"); 
    wb.save(path); wb.close()


def generate(folder, prior_edit=None, raw_rows=None, stale=False):
    folder=Path(folder)
    prior=folder/"prior.xlsx"; net_vendor_prior(prior,stale)
    raw=folder/"raw.xlsx"
    raw_file(raw,raw_rows or [
        ("001111","가상물류","위탁물류비 지급(7월분)",290,0,"2026-09-03"),
        ("001111","가상물류","위탁물류비 지급(8월분)",100,0,"2026-09-04"),
        ("001111","가상물류","9월 청구분",0,50,"2026-09-09"),
    ])
    run=run_reconciliation([prior],[raw],{"25301"},PERIOD)
    out=folder/"out.xlsx"
    write_month_end_statement(out,[prior],run.results,run.new_items,run.issues,PERIOD,
        standalone_debits=run.standalone_debits,journal_items=run.journal_items,
        journal_sources=run.journal_sources,source_paths=[prior,raw])
    return run,out


class OffsetHintTests(unittest.TestCase):
    def test_month_token_needs_exactly_one_month(self):
        self.assertEqual(month_token("위탁물류비 지급(7월분)"),7)
        self.assertEqual(month_token("계정정리)8월 패널티공제"),8)
        self.assertEqual(month_token("12월 청구"),12)
        self.assertIsNone(month_token("9월 본사 임차료 및 8월 관리비"))
        self.assertIsNone(month_token("월말 정산"))
        self.assertIsNone(month_token(""))

    def test_groups_only_months_on_both_sides_and_keep_adjustments(self):
        groups=month_groups(
            [("7월 A",D(100)),("7월 B",D(200)),("계정정리)7월 공제",D(-10)),("8월 C",D(300)),("5월 D",D(9))],
            [("지급(7월분)",D(290)),("지급(8월분)",D(100)),("지급(6월분)",D(1))])
        self.assertEqual(sorted(groups),[7,8])
        self.assertEqual(groups[7].remaining,0)
        self.assertEqual(groups[8].remaining,200)

    def test_subtotal_text_states_the_arithmetic_and_groups(self):
        note=VendorNote(D(590),D(50),D(390),D(250),D(250),"ok",False,
                        (MonthGroup(7,D(290),D(290)),MonthGroup(8,D(300),D(100))))
        text=str(subtotal_note_text(note))
        for expected in ("✓ 검산 일치","전월 합계  590","+ 당월 대변  50","− 당월 차변  390",
                         "= 말 잔액  250","7월 묶음  ✓ 일치 (후보)","청구 290 = 지급 290",
                         "8월 묶음  △ 남음 200 (후보)","청구 300 − 지급 100"):
            self.assertIn(expected,text)

    def test_mismatch_is_never_shown_as_a_pass(self):
        note=VendorNote(D(100),D(0),D(30),D(80),D(70),"mismatch",False)
        text=str(subtotal_note_text(note))
        self.assertIn("✗ 검산 불일치",text)
        self.assertNotIn("✓",text)
        self.assertIn("계산상 잔액  70",text)

    def test_generated_status_words_are_recognised_but_notes_are_not(self):
        for text in ("자동 대사","소계","소계 0","당월 지급 · 7월 묶음 (후보)","자동 대사 ↔ 8행"):
            self.assertTrue(is_generated_status(text),text)
        for text in ("담당자 메모","비고","7월분 확인함",None,123):
            self.assertFalse(is_generated_status(text),text)


class StatementNotesTests(unittest.TestCase):
    def test_net_vendor_rows_and_subtotal_explain_the_offset(self):
        with tempfile.TemporaryDirectory() as folder:
            run,out=generate(folder)
            wb=load_workbook(out,rich_text=True)
            try:
                ws=wb["26.09"]
                rows={r:[ws.cell(r,c).value for c in range(1,7)] for r in range(3,ws.max_row+1)}
                status={r:str(v[5]) for r,v in rows.items() if v[5] is not None}
                joined="\n".join(status.values())
                # claims of a month and payments naming that month carry the same group label
                self.assertEqual(joined.count("순잔액 대사 · 7월 묶음 (후보)"),3)
                self.assertEqual(joined.count("순잔액 대사 · 8월 묶음 (후보)"),1)
                self.assertIn("당월 지급 · 7월 묶음 (후보)\n순잔액에서 차감, 건별 연결 미확정",joined)
                self.assertIn("당월 지급 · 8월 묶음 (후보)",joined)
                subtotal=next(v for r,v in status.items() if "검산" in v)
                for expected in ("✓ 검산 일치","= 말 잔액  250","7월 묶음  ✓ 일치 (후보)",
                                 "8월 묶음  △ 남음 200 (후보)"):
                    self.assertIn(expected,subtotal)
                # the amount column itself is untouched: closing balance matches the subtotal sum
                self.assertEqual(sum(v[4] for v in rows.values() if isinstance(v[4],(int,float))
                                     and v[2]!="소계" and v[0]!="합 계"),250)
            finally:
                wb.close()

    def test_nothing_is_added_below_the_statement_and_columns_A_to_E_are_unchanged(self):
        with tempfile.TemporaryDirectory() as folder:
            run,out=generate(folder)
            wb=load_workbook(out)
            try:
                ws=wb["26.09"]
                # title + header, 4 opening rows + 3 RAW rows (2 payments, 1 credit),
                # 1 vendor subtotal, 1 grand total: no extra rows are ever added.
                self.assertEqual(ws.max_row,2+4+3+1+1)
                self.assertEqual(ws.max_column,6)
                self.assertEqual([c.value for c in ws[ws.max_row]][:1],["합 계"])
                self.assertEqual(wb.sheetnames[0],"26.09")
            finally:
                wb.close()

    def test_matching_decisions_do_not_depend_on_the_new_wording(self):
        with tempfile.TemporaryDirectory() as folder:
            run,_=generate(folder)
            self.assertTrue(all(r.status.name=="SIGNED_NET_AUTO" for r in run.results))


class StaleStatusColumnTests(unittest.TestCase):
    def test_last_months_status_words_in_outside_columns_are_not_carried(self):
        with tempfile.TemporaryDirectory() as folder:
            run,out=generate(folder,stale=True)
            wb=load_workbook(out)
            try:
                ws=wb["26.09"]
                for row in ws.iter_rows(min_row=3):
                    for col in (9,10):       # I and J: last month's leftover status columns
                        cell=row[col-1]
                        self.assertIsNone(cell.value,(cell.coordinate,cell.value))
                        self.assertIsNone(cell.fill.fgColor.rgb if cell.fill.fill_type else None,cell.coordinate)
            finally:
                wb.close()

    def test_owners_own_note_column_is_still_preserved(self):
        with tempfile.TemporaryDirectory() as folder:
            run,out=generate(folder,stale=True)
            wb=load_workbook(out)
            try:
                ws=wb["26.09"]
                self.assertEqual(ws.cell(2,11).value,"비고")
                self.assertIn("담당자 메모",[ws.cell(r,11).value for r in range(3,ws.max_row+1)])
            finally:
                wb.close()


if __name__=="__main__":
    unittest.main()
