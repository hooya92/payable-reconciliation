"""Unsupported files say what to do; the status column stays in F beside the table."""
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from adapters.excel.month_end_statement import write_month_end_statement
from adapters.excel.reader import read_douzone
from application.service import run_reconciliation
from tests.test_statement_end import make_prior, PERIOD
from tests import test_journal_corrections as fixtures


class FileFormatMessageTests(unittest.TestCase):
    def read(self, name, data=b"x"):
        with tempfile.TemporaryDirectory() as folder:
            path=Path(folder)/name
            path.write_bytes(data)
            with self.assertRaises(ValueError) as caught:
                read_douzone(path,{"25301"})
            return str(caught.exception)

    def test_other_formats_are_explained_in_korean_with_what_to_do(self):
        for name in ("a.csv","a.xlsb","a.ods","a.txt"):
            with self.subTest(name):
                message=self.read(name)
                self.assertIn("지원하지 않는 파일 형식",message)
                self.assertIn(".xlsx",message)
                self.assertNotIn("openpyxl",message)

    def test_a_file_that_is_not_really_excel_is_explained(self):
        message=self.read("a.xlsx",b"not a zip")
        self.assertIn("엑셀 파일로 읽을 수 없습니다",message)
        self.assertNotIn("BadZipFile",message)


class StatusColumnTests(unittest.TestCase):
    def status_column(self, tail):
        with tempfile.TemporaryDirectory() as folder:
            prior=make_prior(folder,tail)
            raw_folder=Path(folder)/"raw_dir"; raw_folder.mkdir()
            raw=fixtures.JournalCorrectionTests().make_files(raw_folder,[
                ("001111","가상공영","벨트 교체",700000,0,"2026-09-03")])[1]
            run=run_reconciliation([prior],[raw],{"25301"},PERIOD)
            out=Path(folder)/"out.xlsx"
            write_month_end_statement(out,[prior],run.results,run.new_items,run.issues,PERIOD,
                standalone_debits=run.standalone_debits,journal_items=run.journal_items,
                journal_sources=run.journal_sources,source_paths=[prior,raw])
            ws=load_workbook(out)["26.09"]
            return [c.column_letter for c in ws[2] if c.value=="처리상태"]

    def test_stray_memo_cells_below_the_table_do_not_push_the_status_column_away(self):
        self.assertEqual(self.status_column([(11,9,"메모"),(12,10,"자동 대사"),(14,8,"확인")]),["F"])

    def test_without_any_stray_cells_it_is_f(self):
        self.assertEqual(self.status_column([]),["F"])


if __name__=="__main__":
    unittest.main()


class BrokenXlsAndMergeReasonTests(unittest.TestCase):
    def test_a_broken_xls_says_what_to_do(self):
        message=FileFormatMessageTests().read("a.xls",b"\xd0\xcf\x11\xe0 not really an xls")
        self.assertIn(".xlsx",message)
        self.assertIn("다시 저장",message)

    def test_a_row_that_reads_empty_because_of_a_merge_says_so(self):
        from tests.test_export_tolerance import ExportToleranceTests
        helper=ExportToleranceTests()
        notices,_,_=helper.kept(lambda wb,ws: ws.merge_cells("C3:D3"))
        held=[t for t in notices if "명세서 3행" in t]
        self.assertTrue(held,notices)
        self.assertIn("병합 셀 C3:D3",held[0])
