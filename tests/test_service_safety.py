import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from application.service import run_reconciliation
from adapters.excel.month_end_statement import write_month_end_statement
from domain.models import Status
from domain.period import AccountingPeriod


def make_prior(path):
    wb=Workbook(); ws=wb.active
    ws.append(["거래처코드","거래처명","날짜","적요","금액"])
    ws.append(["001001","가상상사","2026-07-31","7월 유류비",100000])
    wb.save(path)


def make_raw(path, debit=100000, credit=0, desc="7월 유류비"):
    wb=Workbook(); ws=wb.active
    ws.append(["기표일자","계정코드","계정과목명","거래처코드","거래처명","적요","차변","대변"])
    ws.append(["2026-07-05","25301","미지급금-일반","001001","가상상사",desc,debit,credit])
    wb.save(path)


class ServiceSafetyTests(unittest.TestCase):
    def test_overlapping_raw_exports_are_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp)
            prior=root/"prior.xlsx"; raw1=root/"raw_a.xlsx"; raw2=root/"raw_b.xlsx"
            make_prior(prior); make_raw(raw1); make_raw(raw2)
            with self.assertRaisesRegex(ValueError,"Raw 파일 간 동일 전표"):
                run_reconciliation([prior],[raw1,raw2],{"25301"},AccountingPeriod(2026,7))

    def test_orphan_current_debit_is_preserved_as_signed_raw_entry(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp)
            prior=root/"prior.xlsx"; raw=root/"raw.xlsx"
            make_prior(prior)
            wb=Workbook(); ws=wb.active
            ws.append(["기표일자","계정코드","계정과목명","거래처코드",
                       "거래처명","적요","차변","대변"])
            ws.append(["2026-07-05","25301","미지급금-일반",
                       "009999","전월에 없는 업체","미확인 차변",40000,0])
            wb.save(raw)
            result=run_reconciliation([prior],[raw],{"25301"},AccountingPeriod(2026,7))
            self.assertEqual(result.issues,[])
            self.assertEqual(result.results[0].status,Status.UNPAID)
            self.assertEqual(len(result.standalone_debits),1)
            self.assertEqual(result.standalone_debits[0].description,"미확인 차변")
            self.assertEqual(result.standalone_debits[0].debit,40000)

    def test_two_raw_files_keep_source_names_through_export(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp)
            prior=root/"prior.xlsx"; raw1=root/"raw_a.xlsx"; raw2=root/"raw_b.xlsx"
            target=root/"draft.xlsx"
            make_prior(prior)
            for path,when,desc,debit,credit in (
                (raw1,"2026-08-05","7월 유류비",100000,0),
                (raw2,"2026-08-19","8월 새 비용",0,30000),
            ):
                wb=Workbook(); ws=wb.active
                ws.append(["기표일자","계정코드","계정과목명",
                           "거래처코드","거래처명","적요","차변","대변"])
                ws.append([when,"25301","미지급금-일반","001001",
                           "가상상사",desc,debit,credit])
                wb.save(path)
            period=AccountingPeriod(2026,8)
            result=run_reconciliation([prior],[raw1,raw2],{"25301"},period)
            self.assertEqual(len(result.journal_items),2)
            self.assertEqual(
                {result.journal_sources[id(item)] for item in result.journal_items},
                {"raw_a.xlsx","raw_b.xlsx"}
            )
            write_month_end_statement(
                target,[prior],result.results,result.new_items,result.issues,period,
                journal_items=result.journal_items,journal_sources=result.journal_sources
            )
            wb=load_workbook(target,read_only=True)
            try:
                sources={r[6] for r in list(wb["변경내역"].values)[1:]}
                self.assertIn("raw_a.xlsx",sources)
                self.assertIn("raw_b.xlsx",sources)
                self.assertEqual(
                    sum(r[4] for r in list(wb["26.08"].values)[1:]),30000
                )
            finally:
                wb.close()

    def test_blank_account_code_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp)
            prior=root/"prior.xlsx"; raw=root/"raw.xlsx"
            make_prior(prior); make_raw(raw)
            with self.assertRaisesRegex(ValueError,"계정코드"):
                run_reconciliation([prior],[raw],set(),AccountingPeriod(2026,7))


if __name__=="__main__":
    unittest.main()
