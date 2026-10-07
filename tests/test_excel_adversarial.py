import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from adapters.excel.reader import read_douzone, read_prior
from domain.period import AccountingPeriod


DZ_HEADER=["결의일","결의No","순번","기표일자","기표번호","구분","코드","계정과목명","코드","거래처명","적요","차변","대변"]

class ExcelAdversarialTests(unittest.TestCase):
    def _save(self,path,rows,prefix=None):
        wb=Workbook(); ws=wb.active
        for row in (prefix or []): ws.append(row)
        for row in rows: ws.append(row)
        wb.save(path)

    def test_header_row_can_be_offset(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/"raw.xlsx"
            self._save(p,[DZ_HEADER,["2026-08-01",1,1,"2026-08-01",1,"일반","25301","미지급금-일반","051330","가상","A",100,0]], [["보고서 제목"],[],["검색조건"]])
            r=read_douzone(p,{"25301"},AccountingPeriod(2026,8))
            self.assertEqual(len(r.items),1)

    def test_account_filter_excludes_other_accounts(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/"raw.xlsx"
            self._save(p,[DZ_HEADER,
                ["2026-08-01",1,1,"2026-08-01",1,"일반","25301","미지급금-일반","051330","가상","A",100,0],
                ["2026-08-01",2,1,"2026-08-01",2,"일반","99999","다른계정","051330","가상","B",200,0]])
            r=read_douzone(p,{"25301"},AccountingPeriod(2026,8))
            self.assertEqual([x.debit for x in r.items],[100])

    def test_suspicious_amount_is_excluded_and_reported(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/"prior.xlsx"
            self._save(p,[["거래처코드","거래처명","날짜","적요","금액"],["051330","가상","2026-07-31","A","86.000,000원"]])
            r=read_prior(p,"담당자A")
            self.assertEqual(len(r.items),0); self.assertEqual(len(r.issues),1)

    def test_blank_and_subtotal_rows_are_not_items(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/"prior.xlsx"
            self._save(p,[["거래처코드","거래처명","날짜","적요","금액"],
                ["051330","가상","2026-07-31","A",100],[],["","가상 합계","","",100]])
            r=read_prior(p,"담당자A")
            self.assertEqual(len(r.items),1)

    def test_outside_target_month_is_discarded(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/"raw.xlsx"
            self._save(p,[DZ_HEADER,
                ["2026-07-31",1,1,"2026-07-31",1,"일반","25301","미지급금-일반","051330","가상","old",100,0],
                ["2026-08-01",2,1,"2026-08-01",2,"일반","25301","미지급금-일반","051330","가상","target",200,0]])
            r=read_douzone(p,{"25301"},AccountingPeriod(2026,8))
            self.assertEqual(len(r.items),1); self.assertEqual(r.items[0].description,"target")

    def test_missing_required_header_fails_closed(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/"raw.xlsx"
            self._save(p,[["기표일자","거래처명","적요","차변","대변"],["2026-08-01","가상","A",100,0]])
            with self.assertRaises(ValueError): read_douzone(p,{"25301"},AccountingPeriod(2026,8))

if __name__=="__main__":
    unittest.main()
