import tempfile
import unittest
from dataclasses import replace
from decimal import Decimal as D
from pathlib import Path

from openpyxl import load_workbook
from adapters.excel.reader import InputIssue
from adapters.excel.month_end_statement import write_month_end_statement
from application.month_end_plan import StatementRow, build_month_end_plan, issue_detail
from application.service import run_reconciliation
from domain.balance_check import build_balance_checks
from domain.models import JournalLine, Status
from domain.reconciliation import correction_scope
from tests import test_journal_corrections as fixtures

PERIOD=fixtures.PERIOD


def line(desc, debit=0, credit=0, balance=None, row=2, account="25301"):
    return JournalLine("001111","가상공영",account,"미지급금",desc,D(debit),D(credit),
                       "2026-09-01",row,None if balance is None else D(balance))


class ReviewImprovementTests(unittest.TestCase):
    def test_scope_keeps_multiple_same_amount_candidates_and_description_matches(self):
        seed=line("도어_오발행건",credit=-200)
        raws=[seed,line("도어",credit=300),line("다른 비용",debit=200),
              line("벨트",debit=700),line("다른 계정",credit=200,account="25302")]
        raw,opening=correction_scope(raws,[("001111","도어",D(300)),
                                         ("001111","다른 비용",D(200)),("001111","벨트",D(700))])
        self.assertEqual(set(raw),{id(x) for x in raws[:3]})
        self.assertEqual(set(opening),{0,1})
        self.assertIn("연결 후보",opening[0])

    def test_unidentified_correction_does_not_silently_approve_vendor(self):
        seed=line("조정",credit=-50)
        raws=[seed,line("벨트",debit=700)]
        raw,opening=correction_scope(raws,[("001111","도어",D(200))])
        self.assertEqual(len(raw),2)
        self.assertIn("특정할 수 없어",opening[0])

    def test_unrelated_new_credit_of_same_vendor_is_not_blocked(self):
        with tempfile.TemporaryDirectory() as folder:
            helper=fixtures.JournalCorrectionTests()
            run,output=helper.export(folder,[
                ("001111","가상공영","도어 오발행",0,-200000,"2026-09-01"),
                ("001111","가상공영","신규 청소",0,50,"2026-09-03"),
                ("001111","가상공영","벨트 교체",700000,0,"2026-09-05"),
            ])
            self.assertTrue(next(x for x in run.new_items if x.journal.description=="신규 청소").auto_carry)
            wb=load_workbook(output)
            try:
                row=next(r for r in wb["26.09"].iter_rows() if r[3].value=="신규 청소")
                self.assertEqual(row[5].value,"당월 신규")
            finally:
                wb.close()

    def test_input_reason_includes_location_field_and_original_value(self):
        issue=InputIssue("명세서","26.08",11,"금액","17 000?","숫자로 읽을 수 없음","001111")
        text=issue_detail(issue)
        for expected in ("26.08 11행","금액","17 000?","숫자로 읽을 수 없음"):
            self.assertIn(expected,text)

    def test_vendor_difference_explains_both_names_and_source_rows(self):
        with tempfile.TemporaryDirectory() as folder:
            prior,raw=fixtures.JournalCorrectionTests().make_files(folder,[
                ("001111","가상공","벨트 교체",700000,0,"2026-09-01")])
            run=run_reconciliation([prior],[raw],{"25301"},PERIOD)
            reason=run.results[0].reason
            for expected in ("가상공영","가상공","prior.xlsx","raw.xlsx","3행","2행"):
                self.assertIn(expected,reason)

    def balance(self, lines, issues=(), amount=100, output=None, sources=None):
        opening=[("001111","가상공영","전월",D(amount),"2026-08-31",3,{})]
        total=D(amount)+sum((x.credit-x.debit for x in lines),D(0))
        records=[StatementRow("001111","가상공영","합계",total if output is None else D(output),"","",3,{})]
        return build_balance_checks(opening,records,lines,sources or {},issues)[0]

    def test_calculation_and_ledger_checks_are_independent(self):
        missing=self.balance([line("발생",credit=50)])
        self.assertEqual(missing[8],"계산 일치")
        self.assertEqual(missing[11],"비교 미확정")
        self.assertIsNone(missing[9])
        matched=self.balance([line("발생",credit=50,balance=150),line("지급",debit=25,balance=125,row=3)])
        self.assertEqual(matched[6],125)
        self.assertEqual(matched[10],0)
        self.assertEqual(matched[11],"원장 일치")
        different=self.balance([line("발생",credit=50,balance=160)])
        self.assertEqual(different[8],"계산 일치")
        self.assertEqual(different[10],-10)
        self.assertEqual(different[11],"원장 차액 확인")
        dropped=self.balance([line("발생",credit=50)],output=100)
        self.assertEqual(dropped[8],"계산 불일치")

    def test_incomplete_disordered_or_split_ledger_is_not_confirmed(self):
        a=line("발생",credit=50,balance=150)
        b=line("지급",debit=25,balance=125,row=3)
        for lines,sources,issues in (
            ([a,replace(b,ledger_balance=None)],{},[]),
            ([a,replace(b,ledger_balance=D(126))],{},[]),
            ([a,b],{id(a):"a.xlsx",id(b):"b.xlsx"},[]),
            ([a,replace(b,date="2026-08-31")],{},[]),
            ([a,b],{},[InputIssue("더존","Sheet",4,"대변","?","확인")]),
        ):
            with self.subTest(lines=lines,sources=sources):
                self.assertEqual(self.balance(lines,issues,sources=sources)[11],"비교 미확정")

    def test_signed_adjustments_and_multiple_accounts_sum_correctly(self):
        result=self.balance([line("취소",credit=-20,balance=50),
                             line("지급 취소",debit=-10,balance=40,account="25302")])
        self.assertEqual(result[5],90)
        self.assertEqual(result[9],90)
        self.assertEqual(result[11],"원장 일치")

    def test_literal_dash_and_formula_ledger_balances_are_not_confirmed(self):
        from adapters.excel.reader import read_douzone
        with tempfile.TemporaryDirectory() as folder:
            prior,raw=fixtures.JournalCorrectionTests().make_files(folder,[
                ("001111","가상공영","벨트 교체",700000,0,"2026-09-01")])
            for value,expected in (("-",None),("=200000",None),(0,0)):
                wb=load_workbook(raw);ws=wb.active
                ws["I1"]="잔액";ws["I2"]=value
                wb.save(raw);wb.close()
                parsed=read_douzone(raw,{"25301"},PERIOD)
                self.assertEqual(parsed.items[0].ledger_balance,expected)

    def test_raw_balance_export_and_difference_enters_review_sheet(self):
        with tempfile.TemporaryDirectory() as folder:
            prior,raw=fixtures.JournalCorrectionTests().make_files(folder,[
                ("001111","가상공영","벨트 교체",700000,0,"2026-09-01")])
            wb=load_workbook(raw);ws=wb.active
            ws["I1"]="잔액";ws["I2"]=210000
            wb.save(raw);wb.close()
            original=raw.read_bytes()
            run=run_reconciliation([prior],[raw],{"25301"},PERIOD)
            self.assertEqual(run.journal_items[0].ledger_balance,210000)
            output=Path(folder)/"output.xlsx"
            write_month_end_statement(output,[prior],run.results,run.new_items,run.issues,PERIOD,
                journal_items=run.journal_items,journal_sources=run.journal_sources)
            self.assertEqual(raw.read_bytes(),original)
            wb=load_workbook(output)
            try:
                check=list(wb["잔액검증"].values)[1]
                self.assertEqual(check[6],200000)
                self.assertEqual(check[10],-10000)
                self.assertEqual(check[11],"원장 차액 확인")
                self.assertTrue(any(r[0]=="잔액 확인" for r in wb["검토필요"].values))
            finally:
                wb.close()
