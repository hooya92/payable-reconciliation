"""Only settled pairs disappear; every omission remains traceable in 상계내역."""
from pathlib import Path
import tempfile
import unittest

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

from adapters.excel.month_end_statement import write_month_end_statement
from application.service import run_reconciliation
from domain.models import Status
from domain.period import AccountingPeriod


PERIOD=AccountingPeriod(2026,8)


def make_prior(path):
    wb=Workbook()
    ws=wb.active
    ws.title="26.07"
    ws.append(["가상 미지급 명세서"])
    ws.append(["코드","거래처","날짜","적요","금액"])
    for code,name,desc,amount in [
        ("001111","가상AA","유류비AA",1000000),
        ("003333","가상CC","유류비CC",300000),
    ]:
        start=ws.max_row+1
        ws.append(["","","2026-07-31",desc,amount])
        ws.append([code,name,"소계","",f"=SUM(E{start}:E{start})"])
        for cell in ws[ws.max_row][:5]:
            cell.fill=PatternFill("solid",fgColor="FFF2A6")
            cell.font=Font(name="굴림",bold=True)
    wb.save(path)
    wb.close()


def make_raw(path, include_credit=False):
    wb=Workbook()
    ws=wb.active
    ws.title="전표출력"
    ws.append(["기표일자","계정코드","계정과목명","거래처코드",
               "거래처명","적요","차변","대변"])
    ws.append(["2026-08-05","25301","미지급금-일반",
               "001111","가상AA","유류비AA",1000000,0])
    ws.append(["2026-08-12","25301","미지급금-일반",
               "002222","가상BB","유류비BB",500000,0])
    if include_credit:
        ws.append(["2026-08-16","25301","미지급금-일반",
                   "004444","가상DD","8월 청구",0,50000])
        ws.append(["2026-08-20","25301","미지급금-일반",
                   "004444","가상DD","8월 청구",50000,0])
    wb.save(path)
    wb.close()


class MonthEndOffsetAuditTests(unittest.TestCase):
    def _run(self,folder,include_credit=False):
        root=Path(folder)
        prior=root/"opening.xlsx"
        raw=root/"journal.xlsx"
        output=root/"closing.xlsx"
        make_prior(prior)
        make_raw(raw,include_credit=include_credit)
        opening_bytes=prior.read_bytes()
        run=run_reconciliation([prior],[raw],{"25301"},PERIOD)
        write_month_end_statement(
            output,[prior],run.results,run.new_items,run.issues,PERIOD,
            standalone_debits=run.standalone_debits,
            standalone_debit_sources=run.standalone_debit_sources,
            journal_items=run.journal_items,
        )
        self.assertEqual(prior.read_bytes(),opening_bytes)
        return run,output

    def test_exact_aa_pair_removed_bb_negative_shown_and_audit_trace_preserved(self):
        with tempfile.TemporaryDirectory() as folder:
            run,path=self._run(folder)
            self.assertEqual(len(run.standalone_debits),1)
            self.assertEqual(run.issues,[])
            self.assertEqual([r.status for r in run.results],[Status.MATCHED,Status.UNPAID])
            wb=load_workbook(path)
            try:
                ws=wb["26.08"]
                self.assertEqual(ws.max_row,8)
                lines=[(ws.cell(i,4).value,ws.cell(i,5).value,ws.cell(i,3).value)
                       for i in range(3,ws.max_row+1)]
                self.assertEqual([r[0] for r in lines].count("유류비AA"),2)
                self.assertIn(("유류비BB",-500000,"2026-08-12"),lines)
                self.assertIn(("유류비CC",300000,"2026-07-31"),lines)
                self.assertEqual(sum(r[1] for r in lines if r[2]!="소계"),-200000)
                self.assertEqual(ws["E7"].value,"=SUM(E6:E6)")
                self.assertEqual(ws["E7"].fill.fgColor.rgb,"00FFF2A6")
                offsets=list(wb["대사내역"].values)
                self.assertEqual(len(offsets),1)
                self.assertEqual(offsets[1][:5],
                                 ("대사 일치 후보","001111","가상AA","유류비AA",1000000))
                self.assertEqual(offsets[1][8:13],
                                 ("유류비AA","2026-08-05",1000000,None,2))
                changes=[r[0] for r in list(wb["변경내역"].values)[1:]]
                self.assertIn("대사 일치",changes)
                self.assertIn("원장 단독",changes)
                reviews=[r[0] for r in list(wb["검토필요"].values)[1:]]
                self.assertIn("원장 단독",reviews)
            finally:
                wb.close()

    def test_same_month_credit_and_debit_leave_only_audit_not_body(self):
        with tempfile.TemporaryDirectory() as folder:
            run,path=self._run(folder,include_credit=True)
            self.assertEqual(len(run.new_items),1)
            self.assertTrue(run.new_items[0].auto_carry)
            self.assertEqual(run.new_items[0].remaining,0)
            wb=load_workbook(path)
            try:
                entries=[r[3] for r in wb["26.08"].values]
                self.assertEqual(entries.count("8월 청구"),2)
                offsets=[r for r in list(wb["대사내역"].values)[1:]]
                self.assertEqual(len(offsets),2)
                self.assertEqual(offsets[0][0],"대사 일치 후보")
                self.assertEqual(offsets[0][4],1000000)
                self.assertEqual(offsets[0][10:12],(1000000,None))
            finally:
                wb.close()


if __name__=="__main__":
    unittest.main()
