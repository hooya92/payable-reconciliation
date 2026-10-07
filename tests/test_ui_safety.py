import tempfile
import unittest
from pathlib import Path
from collections import Counter

from app import App, changed_snapshot_paths, completion_detail, infer_period_from_inputs, period_alignment_evidence, reconciliation_breakdown, snapshot_file_digests, suggest_reconciliation_period
from domain.models import Status
from domain.period import AccountingPeriod


class UISafetyTests(unittest.TestCase):
    def test_file_attachment_classifier_handler_exists(self):
        self.assertTrue(callable(getattr(App,"_add_classified_files",None)))

    def test_reconciliation_breakdown_separates_auto_review_and_input(self):
        counts=Counter({Status.MATCHED:4,Status.AMBIGUOUS:2,Status.UNPAID:1})
        result=reconciliation_breakdown(counts,new_count=3,issue_count=2)
        self.assertEqual(result,{"matched":4,"review":6,"issues":2,"new":3})

    def test_completion_detail_explicitly_reports_auto_matches(self):
        counts=Counter({Status.MATCHED:4})
        text=completion_detail(AccountingPeriod(2026,8),4,counts)
        self.assertIn("자동 대사 4건",text)
        self.assertIn("2026년 8월",text)

    def test_statement_tabs_can_suggest_target_accounting_month(self):
        current=AccountingPeriod(2026,9)
        suggested=suggest_reconciliation_period(
            current,
            [[AccountingPeriod(2026,6),AccountingPeriod(2026,7)]],
        )
        self.assertEqual(suggested,AccountingPeriod(2026,7))

    def test_latest_statement_month_drives_target_even_if_current_month_also_matches(self):
        current=AccountingPeriod(2026,8)
        suggested=suggest_reconciliation_period(
            current,
            [[AccountingPeriod(2026,6),AccountingPeriod(2026,7),AccountingPeriod(2026,8)]],
        )
        self.assertEqual(suggested,AccountingPeriod(2026,8))

    def test_no_auto_suggestion_when_statement_files_share_no_month(self):
        current=AccountingPeriod(2026,9)
        suggested=suggest_reconciliation_period(
            current,
            [[AccountingPeriod(2026,7)],[AccountingPeriod(2026,8)]],
        )
        self.assertIsNone(suggested)

    def test_raw_month_can_drive_accounting_month_when_statement_has_no_month_tab(self):
        period,source=infer_period_from_inputs(
            [],
            [[AccountingPeriod(2026,7),AccountingPeriod(2026,8)]],
        )
        self.assertEqual(period,AccountingPeriod(2026,8))
        self.assertEqual(source,"raw")

    def test_statement_and_raw_agreement_drives_accounting_month(self):
        period,source=infer_period_from_inputs(
            [[AccountingPeriod(2026,6),AccountingPeriod(2026,7)]],
            [[AccountingPeriod(2026,7),AccountingPeriod(2026,8)]],
        )
        self.assertEqual(period,AccountingPeriod(2026,7))
        self.assertEqual(source,"statement+raw")

    def test_latest_same_statement_raw_month_wins(self):
        period,source=infer_period_from_inputs(
            [[AccountingPeriod(2026,5),AccountingPeriod(2026,6),AccountingPeriod(2026,7)]],
            [[AccountingPeriod(2026,6),AccountingPeriod(2026,7),AccountingPeriod(2026,8)]],
        )
        self.assertEqual(period,AccountingPeriod(2026,7))
        self.assertEqual(source,"statement+raw")

    def test_period_evidence_rejects_coincidental_raw_month_noise(self):
        from openpyxl import Workbook
        with tempfile.TemporaryDirectory() as tmp:
            prior=Path(tmp)/"prior.xlsx"; wb=Workbook(); ws=wb.active; ws.title="26.07"
            ws.append(["거래처코드","거래처명","날짜","적요","금액"])
            ws.append(["001001","가상A","2026-07-31","7월 비용",100]); wb.save(prior)

            raw=Path(tmp)/"raw.xlsx"; wb=Workbook(); ws=wb.active
            ws.append(["기표일자","계정코드","계정과목명","거래처코드","거래처명","적요","차변","대변"])
            ws.append(["2026-07-31","25301","미지급금-일반","099999","기간외","잡음",999,0])
            ws.append(["2026-08-03","25301","미지급금-일반","001001","가상A","7월 비용",100,0])
            wb.save(raw)

            exact,shared=period_alignment_evidence([prior],[raw],{"25301"},AccountingPeriod(2026,7))
            self.assertEqual((exact,shared),(0,0))

    def test_statement_raw_month_conflict_is_not_guessed(self):
        period,source=infer_period_from_inputs(
            [[AccountingPeriod(2026,7)]],
            [[AccountingPeriod(2026,8)]],
        )
        self.assertIsNone(period)
        self.assertEqual(source,"conflict")

    def test_source_snapshot_detects_file_change(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/"input.xlsx"
            p.write_bytes(b"before")
            snapshot=snapshot_file_digests([p])
            self.assertEqual(changed_snapshot_paths(snapshot),[])
            p.write_bytes(b"after")
            self.assertEqual(changed_snapshot_paths(snapshot),["input.xlsx"])


if __name__=="__main__":
    unittest.main()
