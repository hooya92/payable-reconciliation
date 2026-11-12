"""Account codes typed in the UI: one code, several codes, odd separators, typos."""
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from application.service import run_reconciliation
from domain.models import Status
from domain.normalization import parse_account_codes
from domain.period import AccountingPeriod

PERIOD = AccountingPeriod(2026, 9)


def prior(path):
    wb = Workbook(); ws = wb.active; ws.title = "26.08"
    ws.append(["미지급금 명세"]); ws.append(["코드", "거래처", "날짜", "적요", "금액"])
    ws.append([None, None, "2026-08-01", "8월 물류비", 1000000])
    ws.append(["001111", "가상물류", "소계", None, "=SUM(E3:E3)"])
    ws.append([None, None, "2026-08-02", "8월 포장재", 500000])
    ws.append(["002222", "한빛포장", "소계", None, "=SUM(E5:E5)"])
    ws.append(["합 계", None, None, None, '=SUMIF(C3:C6,"소계",E3:E6)'])
    wb.save(path); wb.close()


def raw(path):
    """Two vendors paid under two different payable accounts, plus an unrelated cash line."""
    wb = Workbook(); ws = wb.active
    ws.append(["기표일자", "계정코드", "계정과목명", "거래처코드", "거래처명", "적요", "차변", "대변"])
    ws.append(["2026-09-05", "25301", "미지급금", "001111", "가상물류", "8월 물류비 지급", 1000000, 0])
    ws.append(["2026-09-06", "25302", "외주미지급금", "002222", "한빛포장", "8월 포장재 지급", 500000, 0])
    ws.append(["2026-09-07", "11100", "현금", "009999", "무관", "관계없음", 777, 0])
    wb.save(path); wb.close()


class ParseTests(unittest.TestCase):
    def test_separators(self):
        for text in ("25301,25302", "25301, 25302", "25301，25302", "25301 25302",
                     "25301;25302", "25301\n25302", " 25301 ,, 25302 , "):
            with self.subTest(text=text):
                self.assertEqual(parse_account_codes(text), {"25301", "25302"})

    def test_single_empty_and_none(self):
        self.assertEqual(parse_account_codes("25301"), {"25301"})
        self.assertEqual(parse_account_codes("  ,  "), set())
        self.assertEqual(parse_account_codes(None), set())

    def test_leading_zeros_are_kept(self):
        self.assertEqual(parse_account_codes("0253"), {"0253"})


class SampleRunTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(); self.addCleanup(self.tmp.cleanup)
        d = Path(self.tmp.name)
        self.p = d / "p.xlsx"; self.r = d / "r.xlsx"
        prior(self.p); raw(self.r)

    def run_codes(self, text):
        return run_reconciliation([self.p], [self.r], parse_account_codes(text), PERIOD)

    def matched(self, run):
        return sorted(x.prior.vendor_code for x in run.results if x.status == Status.MATCHED)

    def test_each_single_code_only_reconciles_its_own_account(self):
        self.assertEqual(self.matched(self.run_codes("25301")), ["001111"])
        self.assertEqual(self.matched(self.run_codes("25302")), ["002222"])

    def test_two_codes_reconcile_both_accounts(self):
        for text in ("25301,25302", "25301, 25302", "25301，25302", "25301 25302"):
            with self.subTest(text=text):
                run = self.run_codes(text)
                self.assertEqual(self.matched(run), ["001111", "002222"])
                self.assertEqual(run.unseen_account_codes, [])

    def test_typo_in_one_of_two_codes_is_reported_not_silent(self):
        run = self.run_codes("25301,25399")
        self.assertEqual(self.matched(run), ["001111"])
        self.assertEqual(run.unseen_account_codes, ["25399"])

    def test_unrelated_cash_account_is_never_used(self):
        run = self.run_codes("11100")
        self.assertEqual(self.matched(run), [])

    def test_no_matching_code_names_the_codes_in_the_error(self):
        with self.assertRaises(ValueError) as ctx:
            self.run_codes("25399")
        self.assertIn("25399", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()
