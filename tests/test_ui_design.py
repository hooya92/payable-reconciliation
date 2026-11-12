"""Guards for the design-review fixes in app.py (no display needed)."""
import os
import unittest
from types import SimpleNamespace

import app as A
from domain.period import AccountingPeriod


def luminance(hex_color):
    h=hex_color.lstrip("#")
    r,g,b=[int(h[i:i+2],16)/255 for i in (0,2,4)]
    f=lambda c: c/12.92 if c<=0.03928 else ((c+0.055)/1.055)**2.4
    return 0.2126*f(r)+0.7152*f(g)+0.0722*f(b)


def contrast(fg, bg):
    a,b=sorted((luminance(fg),luminance(bg)),reverse=True)
    return (a+0.05)/(b+0.05)


class ListboxStub:
    def __init__(self, rows, selected=()):
        self.rows=list(rows)
        self.selected=list(selected)

    def curselection(self):
        return tuple(self.selected)

    def delete(self, index, end=None):
        if end=="end":
            self.rows.clear()
        else:
            self.rows.pop(index)


class ContrastTests(unittest.TestCase):
    def test_text_and_status_colours_meet_wcag_aa(self):
        # the pairs the screen really draws
        pairs=[
            (A.TEXT,A.CARD),(A.TEXT,A.BG),(A.TEXT,A.SOFT),(A.TEXT,A.SELECT_BG),
            (A.TEXT2,A.BG),(A.TEXT2,A.CARD),(A.TEXT2,A.SOFT),
            (A.MUTED,A.CARD),(A.MUTED,A.BG),(A.MUTED,A.SOFT),(A.DISABLED_FG,A.SOFT),
            (A.GREEN,A.CARD),(A.GREEN,A.BG),(A.WARN,A.CARD),(A.WARN,A.BG),(A.RED,A.BG),
            (A.ACCENT,A.CARD),(A.ACCENT,A.BG),(A.ACCENT,A.SOFT),("#FFFFFF",A.ACCENT),
        ]
        for fg,bg in pairs:
            with self.subTest(fg=fg,bg=bg):
                self.assertGreaterEqual(contrast(fg,bg),4.5)

    def test_korean_ui_font_is_set_explicitly_on_windows(self):
        self.assertEqual(A.font(12),(A.FONT,12))
        self.assertEqual(A.font(12,True),(A.FONT,12,"bold"))
        self.assertNotIn("Segoe",A.font(12)[0] if os.name=="nt" else "")


class FileRemovalTests(unittest.TestCase):
    def stub(self):
        calls=[]
        return SimpleNamespace(
            _file_add_busy=False,
            _set_banner=lambda *args: calls.append(("banner",args)),
            _after_inputs_removed=lambda: calls.append(("removed",)),
            calls=calls,
        )

    def test_only_selected_files_are_removed(self):
        ui=self.stub()
        paths=["a.xlsx","b.xlsx","c.xlsx"]
        box=ListboxStub(paths,selected=[0,2])
        A.App._remove_selected(ui,box,paths)
        self.assertEqual(paths,["b.xlsx"])
        self.assertEqual(box.rows,["b.xlsx"])
        self.assertIn(("removed",),ui.calls)

    def test_single_file_needs_no_selection(self):
        ui=self.stub()
        paths=["a.xlsx"]
        A.App._remove_selected(ui,ListboxStub(paths),paths)
        self.assertEqual(paths,[])

    def test_several_files_without_selection_asks_instead_of_clearing(self):
        ui=self.stub()
        paths=["a.xlsx","b.xlsx"]
        A.App._remove_selected(ui,ListboxStub(paths),paths)
        self.assertEqual(paths,["a.xlsx","b.xlsx"])
        self.assertEqual(ui.calls[0][0],"banner")

    def test_clear_all_empties_the_list(self):
        ui=self.stub()
        paths=["a.xlsx","b.xlsx"]
        box=ListboxStub(paths)
        A.App._clear_all(ui,box,paths)
        self.assertEqual((paths,box.rows),([],[]))


class PeriodChoiceTests(unittest.TestCase):
    def test_manual_choice_is_kept_and_reported(self):
        pills=[]
        ui=SimpleNamespace(
            year=SimpleNamespace(set=lambda v: setattr(ui,"y",v)),
            month=SimpleNamespace(set=lambda v: setattr(ui,"m",v)),
            year_text=SimpleNamespace(get=lambda: "2026"),
            month_text=SimpleNamespace(get=lambda: "9"),
            _period_manual=False,prior_paths=[],douzone_paths=[],
            _set_detect_pill=pills.append,
            _on_period_changed=lambda: None,
            after_idle=lambda f: None,
        )
        A.App._on_period_menu(ui)
        self.assertEqual((ui.y,ui.m),(2026,9))
        self.assertTrue(ui._period_manual)
        self.assertEqual(pills,["manual"])

    def test_detect_states_never_claim_more_than_is_true(self):
        seen={}
        label=SimpleNamespace(configure=lambda **kw: seen.update(kw))
        ui=SimpleNamespace(detect_pill=label,_sync_period_picker=lambda: None)
        A.App._set_detect_pill(ui,"none")
        self.assertNotIn("자동 설정됨",seen["text"])
        A.App._set_detect_pill(ui,"auto")
        self.assertIn("자동 설정됨",seen["text"])


@unittest.skipUnless(os.environ.get("DISPLAY") or os.name=="nt","needs a display")
class SmokeTests(unittest.TestCase):
    def test_window_builds_with_the_flow_guide_and_run_as_primary_action(self):
        app=A.App()
        try:
            app.geometry("900x740+0+0")
            app.update()
            self.assertEqual(app.run_btn.cget("state"),"normal")
            self.assertEqual(app.month_end_btn.cget("state"),"disabled")
            self.assertEqual(app.save_btn.cget("state"),"disabled")
            # Nothing to list yet: the flow guide is shown instead of an empty table.
            self.assertTrue(app.result_empty.winfo_ismapped())
            self.assertFalse(app.result_main.winfo_ismapped())
            self.assertTrue(app.form.winfo_ismapped())
            self.assertEqual(app.run_btn.cget("fg_color"),A.ACCENT)
            self.assertEqual(app.month_end_btn.cget("fg_color"),A.SOFT)
        finally:
            app.destroy()

    def test_period_pickers_are_a_fallback_shown_only_when_needed(self):
        import tempfile
        from pathlib import Path
        from tests.test_account_codes import prior,raw
        with tempfile.TemporaryDirectory() as d:
            p=Path(d)/"p.xlsx"; r=Path(d)/"r.xlsx"; prior(p); raw(r)
            app=A.App()
            try:
                app.geometry("980x800+0+0")
                app.update()
                # Quiet by default: one line with the month, the pickers hidden.
                self.assertFalse(app.period_picker.winfo_ismapped())
                self.assertEqual(app.period_toggle_btn.cget("text"),"변경")
                self.assertRegex(app.period_value_text.get(),r"^\d{4}년 \d{1,2}월$")
                # The user can open them on purpose, and close them again.
                app._toggle_period_picker(); app.update()
                self.assertTrue(app.period_picker.winfo_ismapped())
                self.assertEqual(app.period_toggle_btn.cget("text"),"접기")
                app._toggle_period_picker(); app.update()
                self.assertFalse(app.period_picker.winfo_ismapped())
                # Detection failed or the files disagree: the user must choose, so they open.
                for state in ("undetected","conflict"):
                    app._set_detect_pill(state); app.update()
                    self.assertTrue(app.period_picker.winfo_ismapped(),state)
                    self.assertFalse(app.period_toggle_btn.winfo_ismapped(),state)
                # Once the situation is over the pickers fold back by themselves (the user never
                # pressed 변경). Files whose month is detected keep them out of the way.
                app.prior_paths.append(str(p)); app.prior_list.insert("end",p.name)
                app.douzone_paths.append(str(r)); app.douzone_list.insert("end",r.name)
                app._maybe_align_period_from_inputs(); app.update()
                self.assertEqual(app._detect_state,"auto")
                self.assertEqual(app.period_value_text.get(),"2026년 9월")
                self.assertFalse(app.period_picker.winfo_ismapped())
            finally:
                app.destroy()

    def test_after_a_run_the_next_step_becomes_primary_and_inputs_collapse(self):
        import tempfile
        import time
        from pathlib import Path
        from tests.test_account_codes import prior,raw
        with tempfile.TemporaryDirectory() as d:
            p=Path(d)/"p.xlsx"; r=Path(d)/"r.xlsx"; prior(p); raw(r)
            app=A.App()
            try:
                app.geometry("900x740+0+0")
                app.prior_paths.append(str(p)); app.prior_list.insert("end",p.name)
                app.douzone_paths.append(str(r)); app.douzone_list.insert("end",r.name)
                app.year.set(2026); app.month.set(9)
                app.account_codes.set("25301，25302")
                app.update()
                app.run()
                deadline=time.time()+30
                while (app._run_busy or app.last_reconciliation_run is None) and time.time()<deadline:
                    app.update(); time.sleep(0.05)
                app.update()
                self.assertIsNotNone(app.last_reconciliation_run)
                self.assertEqual(app.month_end_btn.cget("state"),"normal")
                self.assertEqual(app.month_end_btn.cget("fg_color"),A.ACCENT)
                self.assertEqual(app.run_btn.cget("fg_color"),A.SOFT)
                self.assertTrue(app.result_main.winfo_ismapped())
                self.assertFalse(app.result_empty.winfo_ismapped())
                # Inputs fold into one line; the table keeps usable height at the minimum size.
                self.assertTrue(app.input_bar.winfo_ismapped())
                self.assertFalse(app.form.winfo_ismapped())
                self.assertGreaterEqual(app.detail.winfo_height(),90)
                # Changing an input drops the result and puts the form back.
                app.account_codes.set("25301")
                app.update()
                self.assertTrue(app.form.winfo_ismapped())
                self.assertTrue(app.result_empty.winfo_ismapped())
                self.assertEqual(app.month_end_btn.cget("state"),"disabled")
            finally:
                app.destroy()


if __name__=="__main__":
    unittest.main()
