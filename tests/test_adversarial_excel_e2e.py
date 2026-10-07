import tempfile
import unittest
from collections import Counter
from pathlib import Path

from adapters.excel.reader import read_douzone, read_prior
from domain.models import Status
from domain.period import AccountingPeriod
from domain.reconciliation import new_payables, reconcile
from tests.synthetic_adversarial_factory import build


class AdversarialExcelE2ETests(unittest.TestCase):
    def test_dirty_excel_never_becomes_silent_normal_data(self):
        with tempfile.TemporaryDirectory() as tmp:
            data=build(tmp); expected=data["expected"]

            prior=read_prior(data["prior_path"],"가상담당")
            raw=read_douzone(data["raw_path"],{"25301"},AccountingPeriod(2026,7))
            results=reconcile(prior.items,raw.items)
            counts=Counter(x.status for x in results)

            self.assertEqual(len(prior.issues),expected["PRIOR_INPUT_ISSUES"])
            self.assertEqual(len(raw.issues),expected["RAW_INPUT_ISSUES"])
            self.assertEqual(counts[Status.MATCHED],expected["AUTO_MATCH"])
            self.assertEqual(counts[Status.DESCRIPTION_MISMATCH],expected["DESCRIPTION_MISMATCH"])
            self.assertEqual(counts[Status.VENDOR_NAME_MISMATCH],expected["VENDOR_NAME_MISMATCH"])
            self.assertEqual(
                sum(x.rule=="CODE_AMOUNT_UNIQUE_WITH_NOTE" for x in results),
                expected["AUTO_MATCH_WITH_NOTE"],
            )
            self.assertEqual(counts[Status.VENDOR_MISMATCH],expected["VENDOR_MISMATCH"])
            self.assertEqual(counts[Status.UNPAID],expected["UNPAID"])
            self.assertEqual(
                sum(x.rule=="POSSIBLE_PARTIAL" for x in results),
                expected["POSSIBLE_PARTIAL"],
            )
            self.assertEqual(
                sum(x.rule=="DUPLICATE" for x in results),
                expected["DUPLICATE"],
            )

            fresh=new_payables(raw.items)
            self.assertEqual(len(fresh),expected["CURRENT_MONTH_NEW_REVIEW"])
            self.assertEqual(fresh[0].vendor_code,"002001")


if __name__=="__main__":
    unittest.main()
