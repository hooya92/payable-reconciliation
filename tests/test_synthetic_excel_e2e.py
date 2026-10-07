import tempfile
import unittest
from collections import Counter
from pathlib import Path
from adapters.excel.reader import read_douzone, read_prior
from domain.models import Status
from domain.period import AccountingPeriod
from domain.reconciliation import new_payables, reconcile
from tests.synthetic_excel_factory import build

class SyntheticExcelE2ETests(unittest.TestCase):
    def test_thousands_of_rows_against_golden_answer(self):
        with tempfile.TemporaryDirectory() as tmp:
            expected=build(tmp,prior_count=2400,raw_noise=7200)
            root=Path(tmp); prior_items=[]; issues=[]
            for p in sorted(root.glob("담당자*_미지급.xlsx")):
                r=read_prior(p,p.stem); prior_items.extend(r.items); issues.extend(r.issues)
            dz=read_douzone(root/"더존_Raw_2026_1년.xlsx",{"25301"},AccountingPeriod(2026,8))
            issues.extend(dz.issues)
            result=reconcile(prior_items,dz.items); counts=Counter(x.status for x in result)
            self.assertEqual(len(issues),0)
            self.assertEqual(len(prior_items),expected["prior_count"])
            self.assertEqual(counts[Status.MATCHED],expected["MATCHED"])
            self.assertEqual(counts[Status.DESCRIPTION_MISMATCH],expected["DESCRIPTION_MISMATCH"])
            self.assertEqual(counts[Status.UNPAID],expected["UNPAID"])
            self.assertEqual(counts[Status.VENDOR_MISMATCH],expected["VENDOR_MISMATCH"])
            self.assertEqual(len(new_payables(dz.items)),expected["new_payables"])

if __name__=="__main__":
    unittest.main()
