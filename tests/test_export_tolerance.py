import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from adapters.excel.writer import write_result
from application.month_end_export import export_month_end_draft
from application.service import run_reconciliation
from tests import test_journal_corrections as fixtures

PERIOD=fixtures.PERIOD
ROWS=[("001111","가상공영","벨트 교체",700000,0,"2026-09-03")]


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

    def test_normal_statement_keeps_template(self):
        with tempfile.TemporaryDirectory() as folder:
            prior,raw,run,out=self.prepare(folder)
            kind,reason=export_month_end_draft(out,[prior],run,PERIOD,[prior,raw])
            self.assertEqual((kind,reason),("template",""))
            self.assertIn("26.09",load_workbook(out).sheetnames)

    def test_unkeepable_template_falls_back_to_plain_draft_with_reason(self):
        cases={
            "merged body":lambda wb,ws: ws.merge_cells("C3:D3"),
            "protected sheet":lambda wb,ws: setattr(ws.protection,"sheet",True),
            "existing target sheet":lambda wb,ws: wb.create_sheet("26.09"),
        }
        for name,edit in cases.items():
            with self.subTest(name), tempfile.TemporaryDirectory() as folder:
                prior,raw,run,out=self.prepare(folder,edit)
                kind,reason=export_month_end_draft(out,[prior],run,PERIOD,[prior,raw])
                self.assertEqual(kind,"plain")
                self.assertTrue(reason)
                wb=load_workbook(out)
                try:
                    self.assertIn("당월말명세서 초안",wb.sheetnames)
                    notes=[r for r in wb["요약"].iter_rows(values_only=True) if r[0]=="안내"]
                    self.assertIn(reason,notes[0][1])
                finally:
                    wb.close()

    def test_unreadable_amount_still_reconciles_and_saves(self):
        with tempfile.TemporaryDirectory() as folder:
            prior,raw,run,out=self.prepare(
                folder,lambda wb,ws: ws.__setitem__("E4","N/A"))
            kind,_=export_month_end_draft(out,[prior],run,PERIOD,[prior,raw])
            self.assertEqual(kind,"plain")
            self.assertTrue(run.results)
            self.assertTrue(out.exists())

    def test_never_overwrites_an_input_file(self):
        with tempfile.TemporaryDirectory() as folder:
            prior,raw,run,_=self.prepare(folder,lambda wb,ws: ws.merge_cells("C3:D3"))
            before=prior.read_bytes()
            with self.assertRaises(ValueError):
                export_month_end_draft(prior,[prior],run,PERIOD,[prior,raw])
            self.assertEqual(prior.read_bytes(),before)

    def test_source_text_starting_with_equals_stays_text(self):
        rows=[("001111","가상공영","=1+1",700000,0,"2026-09-03"),
              ("000002","정상 업체","=SUM(",100,0,"2026-09-04")]
        with tempfile.TemporaryDirectory() as folder:
            prior,raw,run,out=self.prepare(folder,rows=rows)
            kind,_=export_month_end_draft(out,[prior],run,PERIOD,[prior,raw])
            self.assertEqual(kind,"template")
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
