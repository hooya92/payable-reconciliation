"""The web screen's logic, driven without any window."""
import tempfile
import time
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from tests.test_account_codes import prior, raw
from webui.session import Session


class FakeDialogs:
    def __init__(self):
        self.to_open = []; self.save_to = None; self.opened = []
    def open_files(self): return list(self.to_open)
    def save_file(self, name): return self.save_to
    def open_folder(self, path): self.opened.append(path)


def wait(session, key="adding", timeout=30):
    end = time.time() + timeout
    while time.time() < end:
        if not getattr(session, key):      # not state(): that would consume the notices
            return
        time.sleep(0.02)
    raise AssertionError("still busy")


class SessionTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(); self.addCleanup(self.tmp.cleanup)
        d = Path(self.tmp.name)
        self.p = d / "p.xlsx"; self.r = d / "r.xlsx"
        prior(self.p); raw(self.r)
        self.dialogs = FakeDialogs()
        self.s = Session(self.dialogs)

    def load(self):
        self.s.add_paths("prior", [self.p]); wait(self.s)
        self.s.add_paths("douzone", [self.r]); wait(self.s)

    def run_it(self):
        self.s.run_reconciliation(); wait(self.s, "running")
        return self.s.state()

    def test_adding_files_detects_the_month_and_names_the_files(self):
        self.s.add_paths("prior", [self.p]); wait(self.s)
        st = self.s.state()
        self.assertEqual(st["files"]["prior"], ["p.xlsx"])
        self.assertEqual(st["period"]["label"], "2026년 9월")
        self.assertEqual(st["period"]["detect"], "auto")
        self.assertFalse(st["period"]["picker_open"])
        self.assertEqual(st["needed_prior"], "2026년 8월")

    def test_a_file_put_in_the_wrong_slot_is_moved_to_the_right_one(self):
        self.s.add_paths("douzone", [self.p]); wait(self.s)        # a statement given as ledger
        st = self.s.state()
        self.assertEqual(st["files"]["prior"], ["p.xlsx"])
        self.assertEqual(st["files"]["douzone"], [])
        self.assertIn("자동 배치", st["banner"]["detail"])

    def test_an_unreadable_file_is_rejected_with_a_reason(self):
        junk = Path(self.tmp.name) / "junk.xlsx"; wb = Workbook(); wb.active.append(["아무 내용"]); wb.save(junk)
        self.s.add_paths("prior", [junk]); wait(self.s)
        st = self.s.state()
        self.assertEqual(st["files"]["prior"], [])
        self.assertEqual(st["notices"][0]["kind"], "warning")
        self.assertIn("junk.xlsx", st["notices"][0]["message"])

    def test_dialog_cancel_changes_nothing(self):
        self.dialogs.to_open = []
        st = self.s.add_files("prior")
        self.assertEqual(st["files"]["prior"], [])
        self.assertFalse(st["busy"]["adding"])

    def test_running_needs_both_inputs(self):
        self.s.add_paths("prior", [self.p]); wait(self.s)
        st = self.s.run_reconciliation()
        self.assertEqual(st["notices"][0]["title"], "파일 필요")
        self.assertFalse(st["done"])

    def test_run_produces_rows_a_summary_and_collapses_the_inputs(self):
        self.load()
        st = self.run_it()
        self.assertTrue(st["done"] and st["can_export"])
        self.assertEqual(st["banner"]["kind"], "success")
        self.assertEqual(len(st["result"]["rows"]), 2)
        row = st["result"]["rows"][0]
        for key in ("vendor", "amount", "status", "matched", "reason", "source"):
            self.assertIn(key, row)
        self.assertIn("명세서 1개", st["recap"])
        self.assertFalse(st["inputs_open"])
        self.assertTrue(st["result"]["preflight"]["ok"])
        self.assertTrue(self.s.open_inputs()["inputs_open"])

    def test_changing_the_inputs_drops_the_result(self):
        self.load(); self.run_it()
        st = self.s.set_account_codes("25301, 25302")
        self.assertFalse(st["done"]); self.assertIsNone(st["result"])
        self.assertEqual(st["banner"]["title"], "입력 조건 변경")
        st = self.run_it()
        self.assertTrue(st["done"])
        st = self.s.remove_file("douzone", 0)
        self.assertIsNone(st["result"]); self.assertEqual(st["files"]["douzone"], [])

    def test_a_month_picked_by_hand_is_kept_and_flagged_when_files_disagree(self):
        self.load()
        st = self.s.set_period(2026, 12)
        self.assertEqual(st["period"]["detect"], "conflict")
        self.assertTrue(st["period"]["forced"] and st["period"]["picker_open"])
        st = self.s.run_reconciliation()
        self.assertEqual(st["banner"]["kind"], "error")
        self.assertEqual(st["notices"][0]["title"], "회계월 확인")

    def test_the_month_picker_opens_on_request(self):
        self.assertFalse(self.s.state()["period"]["picker_open"])
        self.assertTrue(self.s.toggle_picker()["period"]["picker_open"])
        self.assertFalse(self.s.toggle_picker()["period"]["picker_open"])

    def test_month_end_statement_is_saved_where_the_user_chose(self):
        self.load(); self.run_it()
        out = Path(self.tmp.name) / "out.xlsx"
        self.dialogs.save_to = str(out)
        st = self.s.export_month_end()
        self.assertTrue(out.exists())
        self.assertEqual(st["banner"]["title"], "당월 명세서 초안 생성")
        self.assertEqual(self.dialogs.opened, [str(out)])
        wb = load_workbook(out); self.assertIn("검토필요", wb.sheetnames); wb.close()

    def test_cancelling_the_save_dialog_does_nothing(self):
        self.load(); self.run_it()
        self.dialogs.save_to = None
        st = self.s.export_month_end()
        self.assertEqual(st["banner"]["kind"], "success")
        self.assertEqual(st["banner"]["title"], "대사 완료 · 확인 항목 있음" if st["banner"]["title"].endswith("있음") else "대사 완료")

    def test_exporting_after_a_source_file_changed_is_refused(self):
        self.load(); self.run_it()
        self.p.write_bytes(self.p.read_bytes() + b"x")
        self.dialogs.save_to = str(Path(self.tmp.name) / "o.xlsx")
        st = self.s.export_month_end()
        self.assertEqual(st["notices"][0]["title"], "입력 파일 변경")
        self.assertFalse(Path(self.dialogs.save_to).exists())
        self.assertIsNone(st["result"])

    def test_export_before_a_run_is_refused(self):
        st = self.s.export_month_end()
        self.assertEqual(st["notices"][0]["title"], "대사 필요")

    def test_result_file_export(self):
        self.load(); self.run_it()
        out = Path(self.tmp.name) / "result.xlsx"; self.dialogs.save_to = str(out)
        st = self.s.export_result()
        self.assertTrue(out.exists()); self.assertEqual(st["banner"]["title"], "결과 저장 완료")

    def test_notices_are_delivered_once(self):
        self.s.run_reconciliation()
        self.assertEqual(len(self.s.state()["notices"]), 0)   # already consumed by the call above


if __name__ == "__main__":
    unittest.main()
