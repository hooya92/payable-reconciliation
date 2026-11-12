import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from adapters.excel.writer import write_result
from adapters.excel.month_end_statement import write_month_end_statement
from application.service import run_reconciliation
from tests import test_journal_corrections as fixtures

PERIOD=fixtures.PERIOD
ROWS=[("001111","가상공영","벨트 교체",700000,0,"2026-09-03")]


def write_template(out, prior, raw, run):
    """The month-end statement is only ever written in the original template format."""
    write_month_end_statement(
        out,[prior],run.results,run.new_items,run.issues,PERIOD,
        standalone_debits=run.standalone_debits,
        journal_items=run.journal_items,journal_sources=run.journal_sources,
        source_paths=[prior,raw],
    )


def formula_cells(wb, skip=()):
    return [(ws.title,c.coordinate,c.value) for ws in wb if ws.title not in skip
            for row in ws.iter_rows() for c in row if c.data_type=="f"]


class ExportToleranceTests(unittest.TestCase):
    def prepare(self, folder, edit_prior=None, rows=ROWS):
        prior,raw=fixtures.JournalCorrectionTests().make_files(folder,rows)
        if edit_prior:
            wb=load_workbook(prior); edit_prior(wb,wb.active); wb.save(prior); wb.close()
        run=run_reconciliation([prior],[raw],{"25301"},PERIOD)
        return prior,raw,run,Path(folder)/"out.xlsx"

    def reason_for(self, edit):
        """Run the export on a statement that cannot keep its format; return (message, out)."""
        with tempfile.TemporaryDirectory() as folder:
            prior,raw,run,out=self.prepare(folder,edit)
            with self.assertRaises(ValueError) as caught:
                write_template(out,prior,raw,run)
            # No differently shaped file is saved in place of the statement.
            self.assertFalse(out.exists())
            return str(caught.exception)

    def test_normal_statement_keeps_template(self):
        with tempfile.TemporaryDirectory() as folder:
            prior,raw,run,out=self.prepare(folder)
            write_template(out,prior,raw,run)
            self.assertIn("26.09",load_workbook(out).sheetnames)

    def test_unkeepable_template_is_refused_with_a_reason_and_no_other_file(self):
        # Only whole-sheet problems still stop the export. A single unreadable row never does.
        cases={
            "protected sheet":lambda wb,ws: setattr(ws.protection,"sheet",True),
            "existing target sheet":lambda wb,ws: wb.create_sheet("26.09"),
        }
        for name,edit in cases.items():
            with self.subTest(name):
                self.assertTrue(self.reason_for(edit))

    def kept(self, edit):
        """Export a statement with an unreadable part; return (notices, rows of the new sheet)."""
        with tempfile.TemporaryDirectory() as folder:
            prior,raw,run,out=self.prepare(folder,edit)
            notices=write_month_end_statement(
                out,[prior],run.results,run.new_items,run.issues,PERIOD,
                standalone_debits=run.standalone_debits,journal_items=run.journal_items,
                journal_sources=run.journal_sources,source_paths=[prior,raw])
            self.assertTrue(out.exists())
            wb=load_workbook(out)
            try:
                ws=wb["26.09"]
                rows=[[c.value for c in row] for row in ws.iter_rows()]
                merged=[str(r) for r in ws.merged_cells.ranges]
            finally:
                wb.close()
            return notices,rows,merged

    def test_every_unreadable_row_is_kept_and_listed_instead_of_blocking(self):
        def edit(wb,ws):
            ws["E3"]="N/A"; ws["E4"]=None; ws["D6"]=None
        notices,rows,_=self.kept(edit)
        text="\n".join(notices)
        for expected in ("명세서 3행","N/A","명세서 4행","명세서 6행","그대로 두었습니다"):
            self.assertIn(expected,text)
        # the unreadable rows are still on the sheet, flagged, and outside the vendor subtotals
        flat=[cell for row in rows for cell in row]
        self.assertIn("벨트 교체",flat)
        self.assertTrue([c for c in flat if isinstance(c,str) and "읽지 못해 그대로 둠" in c])

    def test_the_four_cases_that_used_to_stop_the_export(self):
        cases={
            # no subtotal row at all
            "no subtotal":(lambda wb,ws:(ws.__setitem__("A5",None),ws.__setitem__("B5",None),
                                         ws.__setitem__("A7",None),ws.__setitem__("B7",None)),"소계"),
            # a subtotal row without its vendor code
            "subtotal without code":(lambda wb,ws: ws.__setitem__("A5",None),"거래처코드"),
            # '-' or blank amount
            "dash and blank amount":(lambda wb,ws:(ws.__setitem__("E3","-"),ws.__setitem__("E4",None)),"명세서 3행"),
            # a merge across transaction cells
            "merged body":(lambda wb,ws: ws.merge_cells("C3:D3"),"C3:D3"),
        }
        for name,(edit,word) in cases.items():
            with self.subTest(name):
                notices,rows,merged=self.kept(edit)
                self.assertTrue([t for t in notices if word in t],notices)
                self.assertEqual(merged,[])
                # the statement is still a statement: header row and a grand-total-less body exist
                self.assertEqual(rows[1][:5],["코드","거래처","날짜","적요","금액"])

    def test_whole_sheet_reasons_name_the_sheet_or_cells_to_change(self):
        cases=(
            (lambda wb,ws: setattr(ws.protection,"sheet",True),("26.08","보호")),
            (lambda wb,ws: wb.create_sheet("26.09"),("26.09","삭제하거나 이름을 바꾼")),
        )
        for edit,words in cases:
            with self.subTest(words):
                reason=self.reason_for(edit)
                for word in words:
                    self.assertIn(word,reason)

    def test_never_overwrites_an_input_file(self):
        with tempfile.TemporaryDirectory() as folder:
            prior,raw,run,_=self.prepare(folder)
            before=prior.read_bytes()
            with self.assertRaises(ValueError):
                write_template(prior,prior,raw,run)
            self.assertEqual(prior.read_bytes(),before)

    def test_source_text_starting_with_equals_stays_text(self):
        rows=[("001111","가상공영","=1+1",700000,0,"2026-09-03"),
              ("000002","정상 업체","=SUM(",100,0,"2026-09-04")]
        with tempfile.TemporaryDirectory() as folder:
            prior,raw,run,out=self.prepare(folder,rows=rows)
            write_template(out,prior,raw,run)
            wb=load_workbook(out)
            try:
                # Only the intended SUM subtotal/total formulas remain on the statement sheet.
                for title,coord,value in formula_cells(wb):
                    self.assertEqual(title,"26.09")
                    self.assertTrue(value.startswith(("=SUM(","=SUMIF(")),(coord,value))
                self.assertEqual(wb["26.09"]["D3"].data_type=="f" or wb["26.09"]["D4"].data_type=="f",False)
            finally:
                wb.close()

    def test_issue_raw_value_formula_is_not_evaluated_in_result_sheet(self):
        with tempfile.TemporaryDirectory() as folder:
            prior,raw,run,_=self.prepare(folder,lambda wb,ws: ws.__setitem__("E4","=SUM("))
            out=Path(folder)/"result.xlsx"
            write_result(out,run.results,run.new_items,run.issues,[prior,raw],PERIOD.label)
            wb=load_workbook(out)
            try:
                self.assertEqual(formula_cells(wb),[])
            finally:
                wb.close()


if __name__=="__main__":
    unittest.main()
