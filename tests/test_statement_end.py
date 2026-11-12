"""The statement ends at its grand-total row; stray cells below it must not block the template."""
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

from adapters.excel.reader import read_prior
from application.month_end_export import export_month_end_draft
from application.service import run_reconciliation
from domain.period import AccountingPeriod
from tests import test_journal_corrections as fixtures

PERIOD=fixtures.PERIOD
YELLOW="FFFF00"
GREEN="A9D08E"


def make_prior(folder, tail=()):
    """Group-subtotal statement ending in a grand total, with formatted empty rows below."""
    wb=Workbook(); ws=wb.active; ws.title="26.08"
    ws.append(["미지급금 명세"]); ws.append(["코드","거래처","날짜","적요","금액"])
    ws.append([None,None,"2026-08-13","벨트 교체",700000])
    ws.append([None,None,"2026-08-31","도어 구매",200000])
    ws.append(["001111","가상공영","소계",None,"=SUM(E3:E4)"])
    ws.append([None,None,"2026-08-31","정상 비용",100])
    ws.append(["000002","정상 업체","소계",None,100])
    ws.append(["합 계",None,None,None,'=SUMIF(C3:C7,"소계",E3:E7)'])
    for row in (5,7):
        for col in range(1,6):
            ws.cell(row,col).fill=PatternFill("solid",fgColor=YELLOW)
    for col in range(1,6):
        ws.cell(8,col).fill=PatternFill("solid",fgColor=GREEN)
    for _ in range(6):
        ws.append([None]*5)
    for row,col,value in tail:
        ws.cell(row,col,value)
    path=Path(folder)/"prior.xlsx"
    wb.save(path); wb.close()
    return path


def export(tail=()):
    with tempfile.TemporaryDirectory() as folder:
        prior=make_prior(folder,tail)
        other=Path(folder)/"raw_dir"; other.mkdir()
        raw=fixtures.JournalCorrectionTests().make_files(other,[
            ("001111","가상공영","벨트 교체",700000,0,"2026-09-03")])[1]
        run=run_reconciliation([prior],[raw],{"25301"},PERIOD)
        out=Path(folder)/"out.xlsx"
        kind,reason=export_month_end_draft(out,[prior],run,PERIOD,[prior,raw])
        if kind!="template":
            return kind,reason,[],{}
        wb=load_workbook(out)
        try:
            rows=[[c.value for c in row] for row in wb["26.09"].iter_rows()]
            fills={i:wb["26.09"].cell(i,1).fill.fgColor.rgb for i in range(1,wb["26.09"].max_row+1)}
        finally:
            wb.close()
        return kind,reason,rows,fills


class StatementEndTests(unittest.TestCase):
    STRAY={
        "space in the amount column":[(11,5," ")],
        "space in the date column":[(11,3," ")],
        "a stray number in the amount column":[(11,5,123)],
        "a stray word in the date column":[(11,3,"확인")],
        "a stray word in the description column":[(11,4,"참고")],
        "whitespace on several rows":[(10,3," "),(12,5,"  "),(14,4," ")],
    }

    def test_stray_cells_below_the_total_keep_the_original_format(self):
        for name,tail in self.STRAY.items():
            with self.subTest(name):
                kind,reason,rows,fills=export([(r,c,v) for r,c,v in tail])
                self.assertEqual((kind,reason),("template",""))
                labels=[row[0] for row in rows]
                self.assertIn("합 계",labels)
                # the original yellow subtotal and green total formatting is carried over
                self.assertIn(YELLOW,{v[-6:] for v in fills.values()})
                self.assertIn(GREEN,{v[-6:] for v in fills.values()})

    def test_stray_value_is_kept_below_the_total_and_not_summed(self):
        kind,_,rows,_=export([(11,5,123)])
        self.assertEqual(kind,"template")
        total_index=next(i for i,row in enumerate(rows) if row[0]=="합 계")
        self.assertIn(123,[row[4] for row in rows[total_index+1:]])
        self.assertNotIn(123,[row[4] for row in rows[:total_index]])
        self.assertTrue(str(rows[total_index][4]).startswith("=SUMIF("))

    def test_a_real_transaction_below_the_total_is_still_reported(self):
        kind,reason,_,_=export([(11,3,"2026-08-31"),(11,4,"추가 비용"),(11,5,500)])
        self.assertEqual(kind,"plain")
        self.assertIn("명세서 11행",reason)
        self.assertIn("소계",reason)

    def test_reader_ignores_stray_cells_below_the_total(self):
        with tempfile.TemporaryDirectory() as folder:
            clean=read_prior(make_prior(folder))
            clean_issues=len(clean.issues)
            clean_items=len(clean.items)+len(clean.review_items)
        with tempfile.TemporaryDirectory() as folder:
            stray=read_prior(make_prior(folder,[(11,5,123),(12,3,"확인"),(13,4," ")]))
            self.assertEqual(len(stray.issues),clean_issues)
            self.assertEqual(len(stray.items)+len(stray.review_items),clean_items)


if __name__=="__main__":
    unittest.main()
