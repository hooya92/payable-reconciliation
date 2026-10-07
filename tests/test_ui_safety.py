import tempfile
import unittest
from pathlib import Path

from app import (
    changed_snapshot_paths,
    snapshot_file_digests,
    statement_confirmation_text,
    validate_statement_confirmation,
)
from domain.period import AccountingPeriod


class UISafetyTests(unittest.TestCase):
    def test_statement_confirmation_uses_previous_month(self):
        period=AccountingPeriod(2026,8)
        self.assertEqual(
            statement_confirmation_text(period),
            "선택한 전월 명세서가 모두 2026년 7월 마감본임을 확인했습니다.",
        )

    def test_statement_confirmation_handles_year_rollover(self):
        period=AccountingPeriod(2026,1)
        self.assertIn("2025년 12월",statement_confirmation_text(period))

    def test_run_is_blocked_without_statement_confirmation(self):
        period=AccountingPeriod(2026,8)
        with self.assertRaisesRegex(ValueError,"2026년 7월"):
            validate_statement_confirmation(False,period)
        validate_statement_confirmation(True,period)

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
