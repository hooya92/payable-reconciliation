"""Regression checks for fast attachment without elapsed-time UI."""
import queue
import unittest

from app import App


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


class AttachmentDisplayTests(unittest.TestCase):
    def test_poll_keeps_pending_name_without_timer_and_finishes(self):
        class Stub:
            def after(self, delay, callback):
                self.poll_delay=delay

            def _poll_file_add(self):
                App._poll_file_add(self)

        obj=Stub()
        obj._file_add_queue=queue.Queue()
        obj._pending_file_labels=["분석 중 · statement.xlsx"]
        obj._pending_file_list=ListboxStub(["existing.xlsx", *obj._pending_file_labels])
        obj._file_add_busy=True
        App._poll_file_add(obj)
        self.assertEqual(obj._pending_file_list.rows,[
            "existing.xlsx","분석 중 · statement.xlsx"
        ])
        self.assertEqual(obj.poll_delay,250)

        obj._file_add_queue.put(([("statement.xlsx","prior",None)],"prior",{},None))
        obj._finish_classified_files=lambda classified,kind,detected: setattr(
            obj,"finished",True
        )
        App._poll_file_add(obj)
        self.assertEqual(obj._pending_file_list.rows,["existing.xlsx"])
        self.assertTrue(obj.finished)
        self.assertFalse(obj._file_add_busy)

    def test_verified_files_show_only_original_filenames(self):
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
            "prior",{},
        )
        self.assertEqual(obj.prior_list.rows,["statement.xlsx"])
        self.assertEqual(obj.douzone_list.rows,["raw.xlsx"])


if __name__=="__main__":
    unittest.main()
