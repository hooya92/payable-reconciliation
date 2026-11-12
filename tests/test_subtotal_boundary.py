"""A subtotal row that cannot be trusted still ends its group: other vendors are unaffected."""
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from adapters.excel.month_end_statement import write_month_end_statement
from adapters.excel.reader import read_prior
from application.service import run_reconciliation
from domain.models import Status
from domain.period import AccountingPeriod

PERIOD = AccountingPeriod(2026, 9)
HEAD = ["기표일자", "계정코드", "계정과목명", "거래처코드", "거래처명", "적요", "차변", "대변"]


def make_prior(path, first_subtotal):
    wb = Workbook(); ws = wb.active; ws.title = "26.08"
    ws.append(["미지급금 명세"]); ws.append(["코드", "거래처", "날짜", "적요", "금액"])
    ws.append([None, None, "2026-08-01", "7월 위탁물류비 GS팀", 1000000])
    ws.append([None, None, "2026-08-01", "8월 위탁물류비 CU팀", 500000])
    ws.append(first_subtotal)                                    # row 5: the broken one
    ws.append([None, None, "2026-08-05", "8월 포장재", 300000])
    ws.append(["002222", "한빛포장", "소계", None, "=SUM(E6:E6)"])
    ws.append(["합 계", None, None, None, '=SUMIF(C3:C7,"소계",E3:E7)'])
    wb.save(path); wb.close()


def make_raw(path):
    wb = Workbook(); ws = wb.active; ws.append(HEAD)
    ws.append(["2026-09-03", "25301", "미지급금", "001111", "가상물류", "8월 위탁물류비 GS팀", 1000000, 0])
    ws.append(["2026-09-04", "25301", "미지급금", "002222", "한빛포장", "8월 포장재", 300000, 0])
    wb.save(path); wb.close()


class SubtotalBoundaryTests(unittest.TestCase):
    BROKEN = {
        "formula, no vendor code": [None, "가상물류", "소계", None, "=SUM(E3:E4)"],
        "number, label in 적요, no code or name": [None, None, None, "소계", 1500000],
        "number, label in 날짜, no code": [None, "가상물류", "소계", None, 1500000],
        "legacy: only the vendor-name column says 소계": [None, "소계", None, None, 1500000],
    }

    def run_case(self, subtotal):
        with tempfile.TemporaryDirectory() as folder:
            prior = Path(folder) / "p.xlsx"; raw = Path(folder) / "r.xlsx"
            make_prior(prior, subtotal); make_raw(raw)
            run = run_reconciliation([prior], [raw], {"25301"}, PERIOD)
            out = Path(folder) / "o.xlsx"
            notices = write_month_end_statement(
                out, [prior], run.results, run.new_items, run.issues, PERIOD,
                standalone_debits=run.standalone_debits, journal_items=run.journal_items,
                journal_sources=run.journal_sources, source_paths=[prior, raw])
            ws = load_workbook(out)["26.09"]
            rows = {r: str(ws.cell(r, 6).value or "") for r in range(3, ws.max_row + 1)
                    if ws.cell(r, 4).value}
            return run, notices, rows

    def test_the_next_vendor_is_still_reconciled(self):
        for name, subtotal in self.BROKEN.items():
            with self.subTest(name):
                run, _, _ = self.run_case(subtotal)
                matched = [x.prior.vendor_code for x in run.results if x.status == Status.MATCHED]
                self.assertEqual(matched, ["002222"], name)

    def test_the_broken_group_is_reported_not_counted_as_a_transaction(self):
        for name, subtotal in self.BROKEN.items():
            with self.subTest(name):
                run, notices, rows = self.run_case(subtotal)
                # the subtotal row is never read as a detail row of anybody
                self.assertFalse([x for x in run.results if x.prior.description == "소계"])
                reported = {issue.row for issue in run.issues}
                self.assertTrue({3, 4} <= reported, (name, reported))
                self.assertTrue([t for t in notices if "3행" in t], notices)
                # the clean vendor's rows pair up
                self.assertTrue([v for v in rows.values() if v.startswith("자동 대사")], rows)

    def test_reader_does_not_carry_the_problem_into_the_next_group(self):
        with tempfile.TemporaryDirectory() as folder:
            prior = Path(folder) / "p.xlsx"
            make_prior(prior, self.BROKEN["formula, no vendor code"])
            read = read_prior(prior, period=PERIOD.previous())
            self.assertEqual([item.vendor_code for item in read.items], ["002222"])
            self.assertFalse([i for i in read.issues if i.row >= 6], read.issues)


if __name__ == "__main__":
    unittest.main()
