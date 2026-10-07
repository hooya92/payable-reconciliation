import tempfile
import unittest
from pathlib import Path
from openpyxl import Workbook
from adapters.excel.reader import read_douzone, read_prior
from domain.period import AccountingPeriod

class ExcelEdgeCaseTests(unittest.TestCase):
    def test_prior_header_offset_subtotal_and_leading_zero_numeric_code(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/"prior.xlsx"; wb=Workbook(); ws=wb.active
            ws.append(["보고서"]); ws.append([""]); ws.append(["거래처코드","거래처명","날짜","적요","금액"])
            ws.append([51330,"가상물류","2026-07-31","7월 운송비",86000000]); ws["A4"].number_format="000000"
            ws.append(["","소계","","",86000000]); wb.save(p)
            r=read_prior(p,"담당자A")
            self.assertEqual(len(r.items),1); self.assertEqual(r.items[0].vendor_code,"051330")
            self.assertEqual(len(r.issues),0)

    def test_douzone_duplicate_code_headers_and_target_month_only(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/"raw.xlsx"; wb=Workbook(); ws=wb.active
            ws.append(["전표출력"]); ws.append(["기표일자","코드","계정과목명","코드","거래처명","적요","차변","대변"])
            ws.append(["2026-07-31",25301,"미지급금-일반",51330,"가상물류","7월",100,0]); ws["D3"].number_format="000000"
            ws.append(["2026-08-01",25301,"미지급금-일반",51330,"가상물류","8월",100,0]); ws["D4"].number_format="000000"
            wb.save(p)
            r=read_douzone(p,{"25301"},AccountingPeriod(2026,8))
            self.assertEqual(len(r.items),1); self.assertEqual(r.items[0].vendor_code,"051330"); self.assertEqual(r.items[0].date,"2026-08-01")

    def test_suspicious_amount_is_excluded_and_reported(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/"prior.xlsx"; wb=Workbook(); ws=wb.active
            ws.append(["거래처코드","거래처명","날짜","적요","금액"]); ws.append(["001234","가상상사","2026-07-31","비용","8600만원"]); wb.save(p)
            r=read_prior(p)
            self.assertEqual(len(r.items),0); self.assertEqual(len(r.issues),1)

    def test_prior_missing_vendor_code_is_reported_not_silently_dropped(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/"prior.xlsx"; wb=Workbook(); ws=wb.active
            ws.append(["거래처코드","거래처명","날짜","적요","금액"])
            ws.append(["","가상상사","2026-07-31","유류비",100000]); wb.save(p)
            r=read_prior(p)
            self.assertEqual(len(r.items),0)
            self.assertTrue(any(x.field=="거래처코드" for x in r.issues))

    def test_prior_negative_amount_is_review_not_normal_data(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/"prior.xlsx"; wb=Workbook(); ws=wb.active
            ws.append(["거래처코드","거래처명","날짜","적요","금액"])
            ws.append(["001234","가상상사","2026-07-31","수정건",-100000]); wb.save(p)
            r=read_prior(p)
            self.assertEqual(len(r.items),0)
            self.assertTrue(any("0 이하" in x.reason for x in r.issues))

    def test_douzone_requires_explicit_account_code(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/"raw.xlsx"; wb=Workbook(); ws=wb.active
            ws.append(["기표일자","계정코드","계정과목명","거래처코드","거래처명","적요","차변","대변"])
            ws.append(["2026-08-01","25301","미지급금-일반","001234","가상상사","유류비",100000,0]); wb.save(p)
            with self.assertRaises(ValueError):
                read_douzone(p,set(),AccountingPeriod(2026,8))

    def test_douzone_missing_account_identifier_fails_closed(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/"raw.xlsx"; wb=Workbook(); ws=wb.active
            ws.append(["기표일자","거래처코드","거래처명","적요","차변","대변"])
            ws.append(["2026-08-01","001234","가상상사","유류비",100000,0]); wb.save(p)
            with self.assertRaises(ValueError):
                read_douzone(p,{"25301"},AccountingPeriod(2026,8))

    def test_douzone_negative_or_two_sided_row_is_review_only(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/"raw.xlsx"; wb=Workbook(); ws=wb.active
            ws.append(["기표일자","계정코드","계정과목명","거래처코드","거래처명","적요","차변","대변"])
            ws.append(["2026-08-01","25301","미지급금-일반","001234","가상상사","역분개",-100000,0])
            ws.append(["2026-08-02","25301","미지급금-일반","001234","가상상사","이상행",100000,100000]); wb.save(p)
            r=read_douzone(p,{"25301"},AccountingPeriod(2026,8))
            self.assertEqual(len(r.items),0)
            self.assertEqual(len(r.issues),2)

    def test_prior_formula_amount_is_quarantined(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/"prior_formula.xlsx"; wb=Workbook(); ws=wb.active
            ws.append(["거래처코드","거래처명","날짜","적요","금액"])
            ws.append(["001234","가상상사","2026-07-31","유류비","=50000+50000"]); wb.save(p)
            r=read_prior(p)
            self.assertEqual(len(r.items),0)
            self.assertTrue(any("수식" in x.reason for x in r.issues))

    def test_douzone_formula_amount_is_quarantined(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/"raw_formula.xlsx"; wb=Workbook(); ws=wb.active
            ws.append(["기표일자","계정코드","계정과목명","거래처코드","거래처명","적요","차변","대변"])
            ws.append(["2026-08-01","25301","미지급금-일반","001234","가상상사","유류비","=50000+50000",0]); wb.save(p)
            r=read_douzone(p,{"25301"},AccountingPeriod(2026,8))
            self.assertEqual(len(r.items),0)
            self.assertTrue(any(x.field=="차변" and "수식" in x.reason for x in r.issues))

if __name__=="__main__": unittest.main()
