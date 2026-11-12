"""A bad ledger row only holds the vendor it names; an unreadable vendor holds everyone."""
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from adapters.excel.month_end_statement import write_month_end_statement
from application.service import run_reconciliation
from domain.period import AccountingPeriod

PERIOD = AccountingPeriod(2026, 9)
HEAD = ["기표일자", "계정코드", "계정과목명", "거래처코드", "거래처명", "적요", "차변", "대변"]


def build(folder, bad_row):
    prior = folder / "prior.xlsx"; raw = folder / "raw.xlsx"
    wb = Workbook(); ws = wb.active; ws.title = "26.08"
    ws.append(["미지급금 명세"]); ws.append(["코드", "거래처", "날짜", "적요", "금액"])
    ws.append([None, None, "2026-08-01", "8월 물류비", 1000000])
    ws.append(["001111", "가상물류", "소계", None, "=SUM(E3:E3)"])
    ws.append([None, None, "2026-08-05", "8월 포장재", 300000])
    ws.append(["002222", "한빛포장", "소계", None, "=SUM(E5:E5)"])
    ws.append(["합 계", None, None, None, '=SUMIF(C3:C6,"소계",E3:E6)'])
    wb.save(prior); wb.close()
    wb = Workbook(); ws = wb.active; ws.append(HEAD)
    ws.append(["2026-09-03", "25301", "미지급금", "001111", "가상물류", "8월 물류비", 1000000, 0])
    ws.append(["2026-09-04", "25301", "미지급금", "002222", "한빛포장", "8월 포장재", 300000, 0])
    ws.append(bad_row)
    wb.save(raw); wb.close()
    run = run_reconciliation([prior], [raw], {"25301"}, PERIOD)
    out = folder / "out.xlsx"
    write_month_end_statement(out, [prior], run.results, run.new_items, run.issues, PERIOD,
        standalone_debits=run.standalone_debits, journal_items=run.journal_items,
        journal_sources=run.journal_sources, source_paths=[prior, raw])
    ws = load_workbook(out)["26.09"]
    notes = {}
    for r in range(3, ws.max_row + 1):
        if ws.cell(r, 3).value == "소계":
            notes[ws.cell(r, 1).value] = str(ws.cell(r, 6).value)
    return run, notes


class LedgerIssueScopeTests(unittest.TestCase):
    def test_a_bad_row_naming_its_vendor_holds_only_that_vendor(self):
        with tempfile.TemporaryDirectory() as d:
            bad = ["2026-09-05", "25301", "미지급금", "001111", "가상물류", "이중 입력", 10, 20]
            run, notes = build(Path(d), bad)
            self.assertEqual(len(run.issues), 1)
            self.assertEqual(run.issues[0].vendor_code, "001111")
            self.assertIn("검산 보류", notes["001111"])
            self.assertIn("✓ 검산 일치", notes["002222"])

    def test_a_bad_row_without_a_readable_vendor_holds_everyone(self):
        with tempfile.TemporaryDirectory() as d:
            bad = ["2026-09-05", "25301", "미지급금", "=1+1", "가상물류", "이중 입력", 10, 0]
            run, notes = build(Path(d), bad)
            self.assertEqual(run.issues[0].vendor_code, "")
            self.assertIn("검산 보류", notes["001111"])
            self.assertIn("검산 보류", notes["002222"])


if __name__ == "__main__":
    unittest.main()
