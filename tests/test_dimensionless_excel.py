"""Regression: valid XLSX exporters may omit the worksheet <dimension> element.

openpyxl in read-only mode then reports max_row/max_column=None even though
streaming iteration and the actual cell values work correctly.
"""
import re
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from openpyxl import Workbook

from adapters.excel.reader import (
    classify_excel_input, detect_douzone_periods, detect_statement_periods,
    read_douzone, read_prior,
)
from domain.period import AccountingPeriod


def omit_sheet_dimensions(path):
    """Simulate the exact metadata omission seen in the 10x test files."""
    source=Path(path)
    repack=source.with_name("repacked.xlsx")
    with ZipFile(source) as src, ZipFile(repack,"w") as dest:
        for member in src.infolist():
            data=src.read(member.filename)
            if member.filename.startswith("xl/worksheets/sheet") and member.filename.endswith(".xml"):
                data=re.sub(rb'<(?:[A-Za-z0-9_]+:)?dimension\s+[^>]*\s*/>',b'',data)
            dest.writestr(member,data)
    repack.replace(source)


class DimensionlessExcelTests(unittest.TestCase):
    def test_douzone_classification_month_and_journal_without_dimensions(self):
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/"sample_10x_august_douzone_raw.xlsx"
            wb=Workbook(); ws=wb.active;ws.title="전표출력"
            ws.append(["결의일","결의No","순번","기표일자","기표번호",
                       "구분","코드","계정과목명","코드","거래처명","적요","차변","대변"])
            ws.append(["2026-08-05",1,1,"2026-08-05",1,"일반",
                       "25301","미지급금-일반","001111","가상물류","7월 운송비",120000,0])
            ws.append(["2026-08-15",2,1,"2026-08-15",2,"일반",
                       "25302","다른계정","009999","제외업체","다른 계정",0,99999])
            wb.save(path);wb.close()
            omit_sheet_dimensions(path)
            self.assertEqual(classify_excel_input(path),"douzone")
            self.assertEqual(detect_douzone_periods(path,{"25301"}),[AccountingPeriod(2026,8)])
            result=read_douzone(path,{"25301"},AccountingPeriod(2026,8))
            self.assertEqual(len(result.items),1)
            self.assertEqual(result.items[0].debit,120000)
            self.assertEqual(result.items[0].vendor_code,"001111")

    def test_prior_statement_works_without_dimensions_and_preserves_group(self):
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/"sample_10x_july_statement.xlsx"
            wb=Workbook();ws=wb.active;ws.title="26.07"
            ws.append(["샘플 미지급 명세서"])
            ws.append([])
            ws.append(["코드","거래처","날짜","적요","금액"])
            ws.append([None,None,"2026-07-31","운송비",120000])
            ws.append([None,None,"2026-07-31","보관료",80000])
            ws.append(["001111","가상물류","소계",None,"=SUM(E4:E5)"])
            wb.save(path);wb.close()
            omit_sheet_dimensions(path)
            self.assertEqual(classify_excel_input(path),"prior")
            self.assertEqual(detect_statement_periods(path),[AccountingPeriod(2026,7)])
            result=read_prior(path,period=AccountingPeriod(2026,7))
            self.assertEqual(len(result.items),2)
            self.assertEqual([item.amount for item in result.items],[120000,80000])
            self.assertEqual(result.issues,[])


if __name__=="__main__":
    unittest.main()
