"""The generated workbook: the statement first, then exactly two companion sheets (검토필요, 대사상세)."""
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from adapters.excel.companion_sheets import (
    DETAIL_MARKER, HEADER_ROW, REVIEW_HEADERS, REVIEW_MARKER, review_data_rows, table_rows,
)
from adapters.excel.month_end_statement import write_month_end_statement
from adapters.excel.reader import read_prior
from application.service import run_reconciliation
from domain.period import AccountingPeriod
from tests import test_journal_corrections as fixtures
from tests.test_month_chain import keep_statement_only, next_month
from tests.test_offset_notes import generate

PERIOD = fixtures.PERIOD


def export(folder, prior, raw):
    run = run_reconciliation([prior], [raw], {"25301"}, PERIOD)
    out = Path(folder) / "out.xlsx"
    notices = write_month_end_statement(out, [prior], run.results, run.new_items, run.issues, PERIOD,
                                        standalone_debits=run.standalone_debits, journal_items=run.journal_items,
                                        journal_sources=run.journal_sources, source_paths=[prior, raw])
    return out, notices


class LayoutTests(unittest.TestCase):
    def test_statement_is_first_and_named_with_its_month(self):
        with tempfile.TemporaryDirectory() as folder:
            _run, out = generate(folder)
            wb = load_workbook(out)
            try:
                self.assertEqual(wb.sheetnames, ["26.09 당월명세서", "검토필요", "대사상세"])
                self.assertIs(wb.active, wb["26.09 당월명세서"])
            finally:
                wb.close()

    def test_companion_sheets_have_a_title_summary_and_styled_header(self):
        with tempfile.TemporaryDirectory() as folder:
            _run, out = generate(folder)
            wb = load_workbook(out)
            try:
                for name in ("검토필요", "대사상세"):
                    ws = wb[name]
                    self.assertTrue(str(ws["A1"].value).startswith(name))
                    self.assertTrue(ws["A2"].value)
                    head = ws.cell(HEADER_ROW, 1)
                    self.assertTrue(head.font.b)
                    self.assertEqual(head.fill.fgColor.rgb, "001F3A5F")
                    self.assertEqual(ws.freeze_panes, f"A{HEADER_ROW + 1}")
                review = wb["검토필요"]
                self.assertEqual(review["J1"].value, REVIEW_MARKER)
                self.assertTrue(review.column_dimensions["J"].hidden)
                self.assertEqual(wb["대사상세"]["L1"].value, DETAIL_MARKER)
                self.assertTrue(wb["대사상세"].column_dimensions["L"].hidden)
                for row in wb["대사상세"].iter_rows(min_row=HEADER_ROW + 1, max_col=5):
                    if isinstance(row[4].value, (int, float)):
                        self.assertEqual(row[4].number_format, '#,##0;[Red]-#,##0;"-"')
            finally:
                wb.close()

    def test_both_sides_of_a_matched_pair_name_each_other(self):
        with tempfile.TemporaryDirectory() as folder:
            prior, raw = fixtures.JournalCorrectionTests().make_files(folder, [
                ("001111", "가상공영", "벨트 교체", 700000, 0, "2026-09-03")])
            out, _ = export(folder, prior, raw)
            wb = load_workbook(out)
            try:
                rows = table_rows(wb["대사상세"])
                prior_side = next(r for r in rows if r[0] == "대사 일치" and r[6] == prior.name)
                raw_side = next(r for r in rows if r[0] == "대사 일치" and r[6] == raw.name)
                self.assertEqual(prior_side[8], f"↔ 더존 {raw_side[7]}행 · 2026-09-03 · 벨트 교체")
                self.assertTrue(raw_side[8].startswith(f"↔ 명세서 {prior_side[7]}행"))
            finally:
                wb.close()

    def test_body_column_follows_the_statement_rows_of_a_pair(self):
        with tempfile.TemporaryDirectory() as folder:
            prior, raw = fixtures.JournalCorrectionTests().make_files(folder, [
                ("001111", "가상공영", "벨트 교체", 700000, 0, "2026-09-03")])
            out, _ = export(folder, prior, raw)
            wb = load_workbook(out)
            try:
                detail = wb["대사상세"]
                self.assertEqual(detail.cell(HEADER_ROW, 10).value, "본문")
                rows = list(detail.iter_rows(min_row=HEADER_ROW + 1))
                prior_side = next(r for r in rows if r[0].value == "대사 일치" and r[6].value == prior.name)
                statement = wb["26.09 당월명세서"]
                mine = prior_side[7].hyperlink.location.split("!")[1].replace("D", "")
                formula = prior_side[9].value
                # both rows of the pair are referenced, and the cell is a live formula (not text)
                self.assertEqual(prior_side[9].data_type, "f")
                self.assertIn(f"'26.09 당월명세서'!$D${mine}", formula)
                self.assertEqual(formula.count("COUNTA("), 4)
                self.assertIn("✓ 짝과 함께 지움", formula)
                carried = next(r for r in rows if r[0].value == "전월 이월")
                self.assertIn("⚠ 지움 (이월·신규 행)", carried[9].value)
                self.assertTrue(str(detail["A2"].value).startswith('="본문 정리'))
                self.assertIsNotNone(statement)
            finally:
                wb.close()

    def test_a_clean_run_says_there_is_nothing_to_review(self):
        with tempfile.TemporaryDirectory() as folder:
            prior, raw = fixtures.JournalCorrectionTests().make_files(folder, [
                ("001111", "가상공영", "벨트 교체", 700000, 0, "2026-09-03")])
            out, _ = export(folder, prior, raw)
            wb = load_workbook(out)
            try:
                ws = wb["검토필요"]
                if review_data_rows(ws) == 0:
                    self.assertIn("0건", ws["A1"].value)
                    self.assertIn("없습니다", ws["A2"].value)
            finally:
                wb.close()


class NextMonthTests(unittest.TestCase):
    def test_reviews_left_in_the_sheet_still_block_next_month(self):
        with tempfile.TemporaryDirectory() as folder:
            prior, raw = fixtures.JournalCorrectionTests().make_files(folder, [
                ("001111", "가상공영", "벨트 교체 1차", 400000, 0, "2026-09-03")])
            out, _ = export(folder, prior, raw)
            wb = load_workbook(out)
            has_reviews = review_data_rows(wb["검토필요"]) > 0
            wb.close()
            self.assertTrue(has_reviews)
            with self.assertRaisesRegex(ValueError, "검토필요"):
                read_prior(out, period=PERIOD)

    def test_an_empty_review_sheet_does_not_block_next_month(self):
        with tempfile.TemporaryDirectory() as folder:
            _run, out = generate(folder)
            wb = load_workbook(out)
            ws = wb["검토필요"]
            ws.delete_rows(HEADER_ROW + 1, ws.max_row)
            wb.save(out); wb.close()
            self.assertEqual(read_prior(out, period=PERIOD).recognized_sheets, ["26.09 당월명세서"])

    def test_old_layout_review_sheet_is_still_recognised(self):
        with tempfile.TemporaryDirectory() as folder:
            _run, out = generate(folder)
            wb = load_workbook(out)
            del wb["검토필요"]
            old = wb.create_sheet("검토필요")
            old.append(REVIEW_HEADERS)
            old["J1"] = REVIEW_MARKER
            old.append(["확인 필요", "001111", "가상물류", "x", 1, "사유", "prior.xlsx", 3])
            wb.save(out); wb.close()
            with self.assertRaisesRegex(ValueError, "검토필요"):
                read_prior(out, period=PERIOD)

    def test_last_months_companion_sheets_are_not_carried_forward(self):
        with tempfile.TemporaryDirectory() as folder:
            _run, out = generate(folder)
            wb = load_workbook(out)
            ws = wb["검토필요"]
            ws.delete_rows(HEADER_ROW + 1, ws.max_row)
            wb.create_sheet("메모")["A1"] = "담당자 메모"
            wb.save(out); wb.close()
            october = next_month(out, folder, [("001111", "가상물류", "10월 청구분", 0, 70, "2026-10-03")], 2026, 10)
            wb = load_workbook(october)
            try:
                self.assertEqual(wb.sheetnames, ["26.10 당월명세서", "메모", "검토필요", "대사상세"])
                self.assertEqual(wb["메모"]["A1"].value, "담당자 메모")
            finally:
                wb.close()

    def test_statement_from_a_multi_month_file_moves_to_the_front(self):
        with tempfile.TemporaryDirectory() as folder:
            prior, raw = fixtures.JournalCorrectionTests().make_files(folder, [
                ("001111", "가상공영", "벨트 교체", 700000, 0, "2026-09-03")])
            wb = load_workbook(prior)
            wb.create_sheet("안내", 0)["A1"] = "참고"
            wb.save(prior); wb.close()
            out, notices = export(folder, prior, raw)
            wb = load_workbook(out)
            try:
                self.assertEqual(wb.sheetnames[:2], ["26.09 당월명세서", "안내"])
                self.assertEqual(sum(ws.sheet_view.tabSelected for ws in wb.worksheets), 1)
            finally:
                wb.close()
            self.assertFalse(any("시트 순서" in n for n in notices))

    def test_a_formula_spanning_sheets_keeps_the_order(self):
        with tempfile.TemporaryDirectory() as folder:
            prior, raw = fixtures.JournalCorrectionTests().make_files(folder, [
                ("001111", "가상공영", "벨트 교체", 700000, 0, "2026-09-03")])
            wb = load_workbook(prior)
            first = wb.create_sheet("집계", 0)
            first["A1"] = "=SUM('26.08:26.08'!E3)"
            wb.save(prior); wb.close()
            out, notices = export(folder, prior, raw)
            wb = load_workbook(out)
            try:
                self.assertEqual(wb.sheetnames[:2], ["집계", "26.09 당월명세서"])
            finally:
                wb.close()
            self.assertTrue(any("시트 순서" in n for n in notices))

    def test_a_template_that_already_has_this_month_is_refused(self):
        with tempfile.TemporaryDirectory() as folder:
            prior, raw = fixtures.JournalCorrectionTests().make_files(folder, [
                ("001111", "가상공영", "벨트 교체", 700000, 0, "2026-09-03")])
            wb = load_workbook(prior)
            wb.create_sheet("26.09 당월명세서")
            wb.save(prior); wb.close()
            with self.assertRaisesRegex(ValueError, "이미 있습니다"):
                export(folder, prior, raw)

    def test_three_months_chain_keeps_reading_the_named_sheet(self):
        with tempfile.TemporaryDirectory() as folder:
            _run, current = generate(folder)
            for year, month, rows in ((2026, 10, [("001111", "가상물류", "10월 청구분", 0, 70, "2026-10-03")]),
                                      (2026, 11, [("001111", "가상물류", "11월 청구분", 0, 80, "2026-11-03")])):
                current = next_month(keep_statement_only(current), folder, rows, year, month)
                self.assertEqual(load_workbook(current, read_only=True).sheetnames[0],
                                 f"{AccountingPeriod(year, month).short_label} 당월명세서")


if __name__ == "__main__":
    unittest.main()
