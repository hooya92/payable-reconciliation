"""Matching must stay roughly linear: a statement of a few thousand rows is a normal month."""
import time
import unittest
from decimal import Decimal as D

from domain.models import JournalLine, PayableItem, SourceRef, Status
from domain.reconciliation import new_payables, reconcile


def build(count):
    items=[]; lines=[]
    for i in range(count):
        code=f"{i // 5:06d}"
        amount=D(100000 + i)
        items.append(PayableItem(code,f"거래처{i // 5}",f"8월 비용{i % 5}",amount,"2026-08-01",i+3,
                                 SourceRef("p.xlsx","26.08",i+3,"")))
        lines.append(JournalLine(code,f"거래처{i // 5}","25301","미지급금",f"8월 비용{i % 5}",amount,D(0),"2026-09-03",i+2))
    return items,lines


class ReconcileScalingTests(unittest.TestCase):
    def test_thousands_of_items_match_quickly_and_correctly(self):
        items,lines=build(4000)
        started=time.perf_counter()
        results=reconcile(items,lines)
        elapsed=time.perf_counter()-started
        self.assertEqual(len(results),4000)
        self.assertTrue(all(r.status==Status.MATCHED for r in results))
        # It was ~10 s when every item rescanned the whole statement (quadratic); now well under 1 s.
        self.assertLess(elapsed,4.0)

    def test_new_payables_does_not_rescan_all_debits_per_credit(self):
        lines=[]
        for i in range(6000):
            code=f"{i % 500:06d}"
            if i % 3==0:
                lines.append(JournalLine(code,"거래처","25301","미지급금",f"7월 비용{i}",D(0),D(100+i),"2026-08-10",i+2))
            else:
                lines.append(JournalLine(code,"거래처","25301","미지급금",f"8월 신규{i}",D(500+i),D(0),"2026-08-10",i+2))
        started=time.perf_counter()
        result=new_payables(lines,[])
        elapsed=time.perf_counter()-started
        self.assertEqual(len(result),2000)   # one per credit line
        self.assertLess(elapsed,3.0)   # ~20 s when every credit rescanned every debit

    def test_a_larger_debit_on_a_competing_item_still_blocks_an_automatic_pairing(self):
        # two opening items of one vendor; the debit equals the small one but the large one
        # could be the one it partly settles, and the descriptions differ -> not auto-assigned
        small=PayableItem("000001","가","A 청구",D(100),"2026-08-01",3,SourceRef("p.xlsx","26.08",3,""))
        large=PayableItem("000001","가","B 청구",D(500),"2026-08-01",4,SourceRef("p.xlsx","26.08",4,""))
        line=JournalLine("000001","가","25301","미지급금","B 청구",D(100),D(0),"2026-09-03",2)
        results={r.prior.description:r for r in reconcile([small,large],[line])}
        self.assertEqual(results["A 청구"].rule,"COMPETING_PARTIAL")


class PeriodCacheTests(unittest.TestCase):
    def test_a_file_is_read_once_until_it_changes(self):
        import os, tempfile
        from unittest import mock
        import app
        with tempfile.TemporaryDirectory() as folder:
            path=os.path.join(folder,"a.xlsx"); open(path,"wb").write(b"x")
            cache={}
            with mock.patch.object(app,"detect_statement_periods",return_value=["m"]) as det:
                self.assertEqual(app.cached_periods(cache,"prior",path,()),["m"])
                app.cached_periods(cache,"prior",path,())
                self.assertEqual(det.call_count,1)
                open(path,"wb").write(b"changed")          # new size -> read again
                app.cached_periods(cache,"prior",path,())
                self.assertEqual(det.call_count,2)
            with mock.patch.object(app,"detect_douzone_periods",return_value=["d"]) as det:
                app.cached_periods(cache,"douzone",path,{"25301"})
                app.cached_periods(cache,"douzone",path,{"25301"})
                app.cached_periods(cache,"douzone",path,{"25301","25302"})   # other codes -> read again
                self.assertEqual(det.call_count,2)


if __name__=="__main__":
    unittest.main()
