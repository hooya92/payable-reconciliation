"""What the page may call. A thin wrapper so pywebview only sees these methods
(it would otherwise walk every attribute of the object it is given)."""
from __future__ import annotations

_PUBLIC = (
    "get_state", "add_files", "add_paths", "remove_file", "clear_files", "set_account_codes",
    "set_period", "toggle_picker", "open_inputs", "run_reconciliation", "export_month_end",
    "export_result", "check_update", "start_update", "minimize_window", "close_window",
)


class Api:
    def __init__(self, session):
        self._session = session

    def get_state(self):
        return self._session.state()

    def add_files(self, kind):
        return self._session.add_files(_kind(kind))

    def add_paths(self, kind, paths):
        return self._session.add_paths(_kind(kind), [str(p) for p in (paths or [])])

    def remove_file(self, kind, index):
        return self._session.remove_file(_kind(kind), int(index))

    def clear_files(self, kind):
        return self._session.clear_files(_kind(kind))

    def set_account_codes(self, text):
        return self._session.set_account_codes(str(text))

    def set_period(self, year, month):
        return self._session.set_period(year, month)

    def toggle_picker(self):
        return self._session.toggle_picker()

    def open_inputs(self):
        return self._session.open_inputs()

    def run_reconciliation(self):
        return self._session.run_reconciliation()

    def export_month_end(self):
        return self._session.export_month_end()

    def export_result(self):
        return self._session.export_result()

    def check_update(self):
        return self._session.check_update()

    def start_update(self):
        return self._session.start_update()

    def minimize_window(self):
        self._session.minimize()

    def close_window(self):
        self._session.quit()


def _kind(kind):
    if kind not in ("prior", "douzone"):
        raise ValueError("알 수 없는 파일 종류입니다.")
    return kind
