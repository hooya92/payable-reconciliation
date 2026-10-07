import tempfile
import unittest
from pathlib import Path

from app import changed_snapshot_paths, snapshot_file_digests


class UISafetyTests(unittest.TestCase):
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
