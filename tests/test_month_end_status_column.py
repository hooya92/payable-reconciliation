"""Preserve original note columns and move analysis states to auxiliary sheets."""
import tempfile
import unittest
from pathlib import Path
from decimal import Decimal

from openpyxl import Workbook,load_workbook
from openpyxl.styles import Font,PatternFill

from adapters.excel.month_end_statement import write_month_end_statement
from domain.models import JournalLine,NewPayable,PayableItem,ReconcileResult,SourceRef,Status
from domain.period import AccountingPeriod

PERIOD=AccountingPeriod(2026,8)

def make_source(path, status_at_f=False):
    wb=Workbook()
    ws=wb.active;ws.title="26.07"
    ws.append(["7월 마감 명세서"])
    ws.append(["코드","거래처","날짜","적요","금액",
               "처리상태" if status_at_f else "담당자 메모","비고" if status_at_f else None])
    ws.append(["","","2026-07-31","유류비",1000000,
               "전월 이월" if status_at_f else "원본 메모","추가 유지" if status_at_f else None])
    ws.append(["001111","가상물류","소계",None,1000000,
               "소계" if status_at_f else "소계 메모","소계 유지" if status_at_f else None])
    for cell in ws[4][:5]:
        cell.fill=PatternFill("solid",fgColor="FFFF99")
        cell.font=Font(name="굴림",bold=True)
    ws["E3"].number_format="#,##0"
    wb.save(path);wb.close()

def opening(path):
    p=PayableItem("001111","가상물류","유류비",Decimal("1000000"),
                  "2026-07-31",3,SourceRef(path.name,"26.07",3,"담당"))
    return [ReconcileResult(Status.UNPAID,p,None,"차변 없음","NO_DEBIT")]

class StatementColumnPreservationTests(unittest.TestCase):
    def test_append_to_rightmost_real_column_without_overwriting_memo(self):
        with tempfile.TemporaryDirectory() as temp:
            p=Path(temp)/"prior.xlsx";o=Path(temp)/"out.xlsx"
            make_source(p)
            write_month_end_statement(o,[p],opening(p),[],[],PERIOD)
            wb=load_workbook(o)
            try:
                ws=wb["26.08"]
                self.assertEqual(ws["F2"].value,"담당자 메모")
                self.assertEqual(ws["F3"].value,"원본 메모")
                self.assertEqual(ws["F4"].value,"소계 메모")
                self.assertEqual(ws["G2"].value,"처리상태")
                self.assertEqual(ws["G3"].value,"전월 이월")
                self.assertEqual(ws["G4"].value,"소계")
                self.assertIn("전월 이월", [x[0] for x in list(wb["변경내역"].values)[1:]])
                self.assertEqual(ws["E3"].number_format,"#,##0")
                self.assertEqual(ws["E4"].fill.fgColor.rgb,"00FFFF99")
            finally:
                wb.close()

    def test_status_moves_right_when_extra_column_added_after_it(self):
        with tempfile.TemporaryDirectory() as temp:
            p=Path(temp)/"prior.xlsx";o=Path(temp)/"out.xlsx"
            make_source(p,status_at_f=True)
            write_month_end_statement(o,[p],opening(p),[],[],PERIOD)
            wb=load_workbook(o)
            try:
                ws=wb["26.08"]
                self.assertEqual(ws["G2"].value,"비고")
                self.assertEqual(ws["G3"].value,"추가 유지")
                self.assertEqual(ws["G4"].value,"소계 유지")
                self.assertIsNone(ws["H2"].value)
                self.assertIsNone(ws["H3"].value)
                self.assertIsNone(ws["H4"].value)
                self.assertEqual(ws["F2"].value,"처리상태")
                self.assertEqual(ws["F3"].value,"전월 이월")
                self.assertEqual(ws["F4"].value,"소계")
            finally:
                wb.close()

    def test_each_row_distinguishes_payment_credit_and_typo_suspect(self):
        with tempfile.TemporaryDirectory() as temp:
            p=Path(temp)/"prior.xlsx";o=Path(temp)/"out.xlsx"
            make_source(p)
            item=opening(p)[0].prior
            debit=JournalLine("001111","가상물류","25301","미지급금",
                              "유류비",Decimal("1000000"),Decimal(0),"2026-08-05",8)
            credit=JournalLine("002222","가상신규","25301","미지급금",
                               "당월 청구",Decimal(0),Decimal("350000"),"2026-08-12",9)
            results=[ReconcileResult(Status.INPUT_TYPO_SUSPECT,item,debit,
                                     "거래처 코드 오타 의심","VENDOR_CODE_TYPO")]
            new=[NewPayable(credit,Decimal("350000"),"신규 확인 필요",False)]
            write_month_end_statement(o,[p],results,new,[],PERIOD,
                                      journal_items=[debit,credit])
            wb=load_workbook(o)
            try:
                ws=wb["26.08"]
                details=[(ws.cell(row,4).value,ws.cell(row,5).value,
                          ws.cell(row,7).value)
                         for row in range(3,ws.max_row+1)
                         if ws.cell(row,3).value!="소계"]
                self.assertEqual(details,[
                    ("유류비",1000000,"오타 의심"),
                    ("유류비",-1000000,"오타 의심"),
                    ("당월 청구",350000,"확인 필요"),
                ])
                states=[r[0] for r in list(wb["검토필요"].values)[1:]]
                self.assertIn("오타 의심",states)
                self.assertIn("확인 필요",states)
                self.assertEqual(ws["G2"].value,"처리상태")
            finally:
                wb.close()

if __name__=="__main__":
    unittest.main()
