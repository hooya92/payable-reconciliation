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
        pairs=[
            (A.TEXT,A.CARD),(A.TEXT2,A.BG),(A.MUTED,A.CARD),(A.MUTED,A.BG),
            (A.MUTED,A.SOFT),(A.MUTED,A.GREEN_BG),(A.MUTED,A.WARN_BG),
            (A.MUTED,A.RED_BG),(A.MUTED,A.RUN_BG),
            (A.GREEN,A.CARD),(A.GREEN,A.GREEN_BG),(A.WARN,A.CARD),(A.WARN,A.WARN_BG),
            (A.RED,A.RED_BG),(A.ACCENT,A.CARD),(A.ACCENT,A.RUN_BG),
            (A.ACCENT,A.ACCENT_SOFT),("#FFFFFF",A.ACCENT),(A.TEXT2,A.SOFT),
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
        ui=SimpleNamespace(detect_pill=label)
        A.App._set_detect_pill(ui,"none")
        self.assertNotIn("자동 설정됨",seen["text"])
        A.App._set_detect_pill(ui,"auto")
        self.assertIn("자동 설정됨",seen["text"])


@unittest.skipUnless(os.environ.get("DISPLAY") or os.name=="nt","needs a display")
class SmokeTests(unittest.TestCase):
    def test_window_builds_and_result_table_is_visible_at_minimum_size(self):
        app=A.App()
        try:
            app.geometry("1080x760+0+0")
            app.update()
            # Not collapsed: header plus part of a row stay visible. Font metrics
            # differ between Windows and Linux, so this is a floor, not an exact size.
            self.assertGreaterEqual(app.detail.winfo_height(),40)
            self.assertEqual(app.run_btn.cget("state"),"normal")
            self.assertEqual(app.month_end_btn.cget("state"),"disabled")
            self.assertEqual(app.save_btn.cget("state"),"disabled")
        finally:
            app.destroy()


if __name__=="__main__":
    unittest.main()
