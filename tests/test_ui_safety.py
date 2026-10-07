import tempfile
import unittest
from pathlib import Path
from collections import Counter

from app import changed_snapshot_paths, completion_detail, reconciliation_breakdown, snapshot_file_digests
from domain.models import Status
from domain.period import AccountingPeriod


class UISafetyTests(unittest.TestCase):
    def test_reconciliation_breakdown_separates_auto_review_and_input(self):
        counts=Counter({Status.MATCHED:4,Status.AMBIGUOUS:2,Status.UNPAID:1})
        result=reconciliation_breakdown(counts,new_count=3,issue_count=2)
        self.assertEqual(result,{"matched":4,"review":6,"issues":2,"new":3})

    def test_completion_detail_explicitly_reports_auto_matches(self):
        counts=Counter({Status.MATCHED:4})
        text=completion_detail(AccountingPeriod(2026,8),4,counts)
        self.assertIn("자동 대사 4건",text)
        self.assertIn("2026년 8월",text)

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
