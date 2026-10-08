import queue
import unittest
from unittest.mock import patch

from app import App, attachment_analysis_label


class ListboxStub:
    def __init__(self, rows=()):
        self.rows=list(rows)

    def get(self, start, end):
        return tuple(self.rows)

    def delete(self, index, end=None):
        if end == "end":
            self.rows.clear()
        else:
            self.rows.pop(index)

    def insert(self, index, value):
        if index == "end":
            self.rows.append(value)
        else:
            self.rows.insert(index, value)


class AttachmentProgressTests(unittest.TestCase):
    def test_labels_show_elapsed_and_completed_duration(self):
        self.assertEqual(
            attachment_analysis_label("statement.xlsx", 65),
            "분석 중 (01:05) · statement.xlsx",
        )
        self.assertEqual(
            attachment_analysis_label("statement.xlsx", 65, completed=True),
            "분석 완료 (01:05) · statement.xlsx",
        )

    def test_poll_updates_pending_list_and_warns_after_thirty_seconds(self):
        class Stub:
            def _set_banner(self, kind, title, detail):
                self.banner=(kind,title,detail)

            def after(self, delay, callback):
                self.poll_delay=delay

        obj=Stub()
        obj._file_add_queue=queue.Queue()
        obj._file_add_started=100
        obj._file_add_stage="명세서 회계월 감지"
        obj._pending_file_paths=("statement.xlsx",)
        obj._pending_file_labels=[attachment_analysis_label("statement.xlsx", 0)]
        obj._pending_file_list=ListboxStub(["original.xlsx", *obj._pending_file_labels])
        with patch("app.time.monotonic", return_value=135):
            App._poll_file_add(obj)
        self.assertEqual(obj._pending_file_list.rows, [
            "original.xlsx", "분석 중 (00:35) · statement.xlsx"
        ])
        self.assertIn("30초 이상",obj.banner[2])
        self.assertEqual(obj.poll_delay,250)

        obj._file_add_queue.put(([("statement.xlsx","prior",None)],"prior",{},None))
        obj._finish_classified_files=lambda classified, kind, detected, seconds: setattr(
            obj,"finished_seconds",seconds
        )
        with patch("app.time.monotonic", return_value=142):
            App._poll_file_add(obj)
        self.assertEqual(obj._pending_file_list.rows, ["original.xlsx"])
        self.assertEqual(obj.finished_seconds,42)
        self.assertFalse(obj._file_add_busy)

    def test_verified_files_keep_completed_label(self):
        class Stub:
            def _refresh_file_counts(self): pass
            def _invalidate_results(self): pass
            def _maybe_align_period_from_inputs(self, detected=None): pass
            def _set_banner(self, kind, title, detail): pass

        obj=Stub()
        obj.prior_paths=[]
        obj.douzone_paths=[]
        obj.prior_list=ListboxStub()
        obj.douzone_list=ListboxStub()
        App._finish_classified_files(
            obj,[("statement.xlsx","prior",None),("raw.xlsx","douzone",None)],
            "prior",{},19,
        )
        self.assertEqual(obj.prior_list.rows, ["분석 완료 (00:19) · statement.xlsx"])
        self.assertEqual(obj.douzone_list.rows, ["분석 완료 (00:19) · raw.xlsx"])


if __name__ == "__main__":
    unittest.main()
