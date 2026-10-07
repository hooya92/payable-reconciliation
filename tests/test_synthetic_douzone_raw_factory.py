import tempfile
import unittest

from adapters.excel.reader import read_douzone
from domain.period import AccountingPeriod
from tests.synthetic_douzone_raw_factory import build


class CleanDouzoneRawFactoryTests(unittest.TestCase):
    def test_generated_raw_matches_expected_douzone_shape_and_period_filter(self):
        with tempfile.TemporaryDirectory() as tmp:
            path=build(tmp)
            result=read_douzone(path,{"25301"},AccountingPeriod(2026,8))

            self.assertEqual(len(result.issues),0)
            self.assertEqual(len(result.items),5)
            self.assertTrue(all(x.account_code=="25301" for x in result.items))
            self.assertTrue(all(x.date.startswith("2026-08") for x in result.items))

            debits=[x for x in result.items if x.debit>0]
            credits=[x for x in result.items if x.credit>0]
            self.assertEqual(len(debits),2)
            self.assertEqual(len(credits),3)
            self.assertEqual(debits[0].vendor_code,"051330")


if __name__=="__main__":
    unittest.main()
