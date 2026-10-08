import tempfile
import unittest

from application.service import run_reconciliation
from domain.models import Status
from domain.period import AccountingPeriod
from tests.synthetic_clean_factory import build


class SyntheticCleanFactoryTests(unittest.TestCase):
    def test_clean_real_like_dataset_is_all_auto_matched(self):
        with tempfile.TemporaryDirectory() as tmp:
            data=build(tmp)
            run=run_reconciliation(
                [data["prior_path"]],
                [data["raw_path"]],
                {"25301"},
                AccountingPeriod(2026,8),
            )
            self.assertEqual(run.prior_count,4)
            self.assertEqual(run.counts[Status.MATCHED],4)
            self.assertEqual(len(run.issues),0)
            self.assertEqual(len(run.new_items),0)
            self.assertTrue(all(x.status==Status.MATCHED for x in run.results))


if __name__=="__main__":
    unittest.main()
