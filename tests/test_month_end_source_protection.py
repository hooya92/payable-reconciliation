import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from app import App, open_saved_folder, snapshot_file_digests
from adapters.excel.month_end_statement import write_month_end_statement
from domain.period import AccountingPeriod
from tests.test_month_end_template import write_flat


class MonthEndSourceProtectionTests(unittest.TestCase):
    def test_permission_error_shows_recovery_guidance_and_allows_retry(self):
        run=SimpleNamespace(results=[],new_items=[],issues=[],
                            standalone_debits=[],standalone_debit_sources={},
                            journal_items=[],journal_sources={})
        ui=SimpleNamespace(last_period=AccountingPeriod(2026,8),
                           last_reconciliation_run=run,last_source_digests={},
                           last_source_paths=[],prior_paths=[],
                           _set_banner=lambda *args: None)
        with patch("app.filedialog.asksaveasfilename",return_value="result.xlsx"), \
                patch("app.write_month_end_statement",
                      side_effect=[PermissionError(13,"Permission denied","result.xlsx"),None]) as writer, \
                patch("app.messagebox.showerror") as error, \
                patch("app.open_saved_folder") as reveal:
            App.export_month_end(ui)
            reveal.assert_not_called()
            error.assert_called_once()
            title,message=error.call_args.args
            self.assertEqual(title,"파일 접근 거부")
            for text in ("Excel","닫은 뒤","다른 파일명","쓰기 권한","result.xlsx"):
                self.assertIn(text,message)
            self.assertIs(ui.last_reconciliation_run,run)
            App.export_month_end(ui)
            self.assertEqual(writer.call_count,2)
            self.assertEqual(error.call_count,1)
            reveal.assert_called_once_with("result.xlsx")

    def test_windows_opens_only_parent_folder(self):
        with tempfile.TemporaryDirectory() as folder:
            output=Path(folder)/"명세서 결과.xlsx"
            with patch("app.sys.platform","win32"), \
                    patch("app.os.startfile",create=True) as startfile:
                open_saved_folder(output)
            startfile.assert_called_once_with(str(output.resolve().parent))

    def test_folder_open_failure_preserves_saved_result(self):
        with tempfile.TemporaryDirectory() as folder:
            prior=Path(folder)/"prior.xlsx"
            output=Path(folder)/"result.xlsx"
            write_flat(prior)
            run=SimpleNamespace(results=[],new_items=[],issues=[],
                                standalone_debits=[],standalone_debit_sources={},
                                journal_items=[],journal_sources={})
            banners=[]
            ui=SimpleNamespace(last_period=AccountingPeriod(2026,8),
                               last_reconciliation_run=run,
                               last_source_digests=snapshot_file_digests([prior]),
                               last_source_paths=[prior],prior_paths=[prior],
                               _set_banner=lambda *args: banners.append(args))
            with patch("app.filedialog.asksaveasfilename",return_value=str(output)), \
                    patch("app.open_saved_folder",side_effect=PermissionError(13,"Denied")), \
                    patch("app.messagebox.showerror") as error, \
                    patch("app.messagebox.showwarning") as warning:
                App.export_month_end(ui)
            self.assertTrue(output.is_file())
            self.assertEqual(banners[-1][0],"success")
            error.assert_not_called()
            warning.assert_called_once()
            self.assertIn("정상적으로 저장",warning.call_args.args[1])
            self.assertIn(str(output.resolve()),warning.call_args.args[1])

    def test_cancelled_save_does_not_open_folder(self):
        run=SimpleNamespace()
        ui=SimpleNamespace(last_period=AccountingPeriod(2026,8),
                           last_reconciliation_run=run,last_source_digests={})
        with patch("app.filedialog.asksaveasfilename",return_value=""), \
                patch("app.open_saved_folder") as reveal, \
                patch("app.write_month_end_statement") as writer:
            App.export_month_end(ui)
        reveal.assert_not_called()
        writer.assert_not_called()

    def test_ui_rejects_raw_output_without_changing_inputs(self):
        with tempfile.TemporaryDirectory() as folder:
            prior = Path(folder) / "prior.xlsx"
            raw = Path(folder) / "raw.xlsx"
            write_flat(prior)
            write_flat(raw)
            sources = [prior, raw]
            before = snapshot_file_digests(sources)
            run = SimpleNamespace(results=[], new_items=[], issues=[],
                                  standalone_debits=[], standalone_debit_sources={},
                                  journal_items=[], journal_sources={})
            ui = SimpleNamespace(last_period=AccountingPeriod(2026, 8),
                                 last_reconciliation_run=run,
                                 last_source_digests=before, last_source_paths=sources,
                                 prior_paths=[prior], _set_banner=lambda *args: None)
            with patch("app.filedialog.asksaveasfilename", return_value=str(raw)), \
                    patch("app.messagebox.showerror") as error:
                App.export_month_end(ui)
            error.assert_called_once()
            self.assertIn("원본", error.call_args.args[1])
            self.assertEqual(snapshot_file_digests(sources), before)

    def test_writer_rejects_raw_and_alias_paths(self):
        with tempfile.TemporaryDirectory() as folder:
            prior = Path(folder) / "prior.xlsx"
            raw = Path(folder) / "raw.xlsx"
            write_flat(prior)
            write_flat(raw)
            alias = Path(folder) / "alias.xlsx"
            alias.symlink_to(raw)
            before = snapshot_file_digests([prior, raw])
            for output in (raw, alias):
                with self.subTest(output=output):
                    with self.assertRaisesRegex(ValueError, "원본"):
                        write_month_end_statement(
                            output, [prior], [], [], [], AccountingPeriod(2026, 8),
                            source_paths=[prior, raw],
                        )
                    self.assertEqual(snapshot_file_digests([prior, raw]), before)
