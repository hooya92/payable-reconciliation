from __future__ import annotations
import os, tempfile, time, unittest
from pathlib import Path
from adapters.excel.reader import read_douzone
from domain.period import AccountingPeriod
from tests.synthetic_excel_factory import build_synthetic_dataset

@unittest.skipUnless(os.getenv("RUN_STRESS") == "1", "manual stress test")
class LargeRawStressTests(unittest.TestCase):
    def test_large_raw_selected_month_extraction(self):
        rows=int(os.getenv("STRESS_RAW_NOISE","50000"))
        with tempfile.TemporaryDirectory() as tmp:
            data=build_synthetic_dataset(Path(tmp), prior_count=2400, raw_noise=rows, new_count=600)
            start=time.perf_counter()
            result=read_douzone(data["raw_path"], {"25301"}, AccountingPeriod(2026,8))
            elapsed=time.perf_counter()-start
            self.assertGreater(len(result.items),0)
            print(f"STRESS raw_noise={rows:,} selected={len(result.items):,} elapsed={elapsed:.2f}s")

if __name__=="__main__": unittest.main()
