"""The screen's state and actions, with no GUI toolkit in sight.

The page asks `state()` for everything it draws and calls the action methods; the same object
is driven by pywebview in the app, by a tiny HTTP bridge in development, and by plain unit tests.
Slow work (reading Excel, matching, saving) runs in worker threads; `state()['busy']` tells
the page to keep polling.
"""
from __future__ import annotations

import threading
from datetime import date
from pathlib import Path

from adapters import updater
from adapters.excel.month_end_statement import write_month_end_statement
from adapters.excel.reader import classify_excel_input
from adapters.excel.writer import write_result
from application.service import run_reconciliation
from application.workflow import (
    cached_periods, changed_snapshot_paths, completion_detail, find_duplicate_files,
    infer_period_from_inputs, open_saved_folder, reconciliation_breakdown, snapshot_file_digests,
)
from domain.models import Status
from domain.normalization import parse_account_codes
from domain.period import AccountingPeriod

IDLE_HINT = "파일 추가 → 대사 시작 → 결과 검토 후 '당월 명세서 생성'"
AUTO = (Status.MATCHED, Status.SIGNED_NET_AUTO)
DETECT_TEXT = {
    "none": "파일을 추가하면 자동 감지",
    "auto": "파일에서 자동 설정됨",
    "manual": "직접 선택함",
    "undetected": "월을 감지하지 못함 · 직접 선택",
    "conflict": "파일의 월과 다름",
}


class Session:
    def __init__(self, dialogs, quit_callback=None, today=None):
        self.dialogs = dialogs            # open_files() / save_file(name) / open_folder(path)
        self.quit = quit_callback or (lambda: None)
        self.lock = threading.RLock()
        today = today or date.today()
        self.year, self.month = today.year, today.month
        self.prior_paths, self.douzone_paths = [], []
        self.account_codes = "25301"
        self.period_manual = False
        self.detect_state = "none"
        self.picker_open = False
        self.banner = {"kind": "idle", "title": "대사 전", "detail": IDLE_HINT}
        self.notices = []
        self.pending = []
        self.adding = False
        self.running = False
        self.inputs_open = False
        self.period_cache = {}
        self._reset_results()
        self.update_release = None
        self.update_state = {"busy": False, "percent": 0, "error": ""}

    # ------------------------------------------------------------------ state
    def _reset_results(self):
        self.results, self.new_items, self.issues = [], [], []
        self.run = None
        self.last_period = None
        self.last_source_paths, self.last_source_digests = [], {}
        self.rows, self.counts, self.preflight = [], {}, {"text": "", "ok": True}

    def _period(self):
        try:
            return AccountingPeriod(int(self.year), int(self.month))
        except Exception:
            return None

    def state(self):
        with self.lock:
            p = self._period()
            forced = self.detect_state in ("undetected", "conflict")
            done = self.run is not None and not self.running
            codes = ", ".join(sorted(parse_account_codes(self.account_codes))) or "-"
            today = date.today()
            data = {
                "version": updater.version_label(),
                "busy": {"adding": self.adding, "running": self.running},
                "period": {
                    "year": self.year, "month": self.month,
                    "label": f"{p.year}년 {p.month}월" if p else "-",
                    "summary": (f"{p.previous().short_label} 말 명세서 + {p.short_label} 원장 → {p.short_label} 말 초안"
                                if p else "파일을 추가하면 회계월을 자동으로 감지합니다."),
                    "detect": self.detect_state, "detect_text": DETECT_TEXT[self.detect_state],
                    "forced": forced, "picker_open": forced or self.picker_open,
                    "years": list(range(today.year - 10, today.year + 2)),
                },
                "files": {
                    "prior": [Path(x).name for x in self.prior_paths],
                    "douzone": [Path(x).name for x in self.douzone_paths],
                    "pending": list(self.pending),
                },
                "needed_prior": p.previous().label if p else "",
                "account_codes": self.account_codes,
                "banner": dict(self.banner),
                "can_run": not self.running and not self.adding,
                "can_export": done,
                "done": done,
                "inputs_open": self.inputs_open,
                "recap": (f"명세서 {len(self.prior_paths)}개 · 더존 원장 {len(self.douzone_paths)}개"
                          f" · 계정코드 {codes} · {self.last_period.short_label if self.last_period else ''} 대상"),
                "result": ({"summary": self._summary(), "rows": self.rows, "preflight": self.preflight}
                           if self.run is not None else None),
                "update": self._update_view(),
                "notices": self.notices,
            }
            self.notices = []
            return data

    def _summary(self):
        b = reconciliation_breakdown(self.counts, len(self.new_items), len(self.issues))
        return {"matched": b["matched"], "review": b["review"], "issues": b["issues"], "new": b["new"]}

    def _banner(self, kind, title, detail):
        self.banner = {"kind": kind, "title": title, "detail": detail}

    def _notice(self, kind, title, message):
        self.notices.append({"kind": kind, "title": title, "message": message})

    # ------------------------------------------------------------ period logic
    def _set_detect(self, state):
        self.detect_state = state

    def _detected(self, kind, path, codes):
        try:
            return cached_periods(self.period_cache, kind, path, codes)
        except Exception:
            return []

    def _align_period(self, show_banner=True, detected=None):
        """Find the target month from the files. True/False/None as for the Tk screen:
        True = settled, False = the files disagree, None = nothing to go on."""
        if not self.prior_paths and not self.douzone_paths:
            self._set_detect("manual" if self.period_manual else "none")
            if show_banner:
                self._banner("idle", "대사 전", IDLE_HINT)
            return None
        codes = parse_account_codes(self.account_codes)
        statement_groups, statement_labels, raw_groups, raw_labels = [], [], [], []
        for path in self.prior_paths:
            periods = detected[path] if detected and path in detected else self._detected("prior", path, ())
            if periods:
                statement_groups.append(periods); statement_labels.extend(periods)
        for path in self.douzone_paths:
            periods = detected[path] if detected and path in detected else self._detected("douzone", path, codes)
            if periods:
                raw_groups.append(periods); raw_labels.extend(periods)
        current = self._period()
        suggested, source = infer_period_from_inputs(statement_groups, raw_groups, preferred=current)
        if suggested is None:
            if source == "conflict":
                self._set_detect("conflict")
                if show_banner:
                    st = ", ".join(p.label for p in sorted(set(statement_labels))) or "감지 안 됨"
                    rt = ", ".join(p.label for p in sorted(set(raw_labels))) or "감지 안 됨"
                    self._banner("warning", "회계월 자동 설정 보류",
                                 f"전월 명세서({st})와 당월 더존 원장({rt})가 이어지지 않습니다.")
                return False
            self._set_detect("manual" if self.period_manual else "undetected")
            return None
        if suggested != current:
            if self.period_manual:
                self._set_detect("conflict")
                if show_banner:
                    self._banner("warning", "선택한 회계월 확인",
                                 f"선택한 {current.label}과 파일에서 감지한 {suggested.label}이 다릅니다. "
                                 "월을 다시 선택하거나 파일을 확인해주세요.")
                return False
            self.year, self.month = suggested.year, suggested.month
            self._invalidate()
        self._set_detect("manual" if self.period_manual else "auto")
        if show_banner:
            if source == "statement+raw":
                detail = f"{suggested.previous().label} 명세서 + {suggested.label} 더존 원장를 확인해 대상 회계월을 자동 설정했습니다."
            elif source == "statement":
                detail = f"전월 명세서 기준으로 당월 회계월을 {suggested.label}로 자동 설정했습니다."
            else:
                detail = f"더존 원장의 미지급금 전표월을 기준으로 대상 회계월을 {suggested.label}로 자동 설정했습니다."
            self._banner("idle", "회계월 자동 설정", detail)
        return True

    def _invalidate(self):
        had = bool(self.results or self.new_items or self.issues or self.last_period)
        self._reset_results()
        self.inputs_open = False
        if had:
            self._banner("warning", "입력 조건 변경", "입력 조건이 변경되었습니다. 다시 대사를 실행하세요.")
        else:
            self._banner("idle", "대사 전", IDLE_HINT)

    # ----------------------------------------------------------- input actions
    def set_period(self, year, month):
        with self.lock:
            try:
                AccountingPeriod(int(year), int(month))
            except Exception:
                return self.state()
            self.year, self.month = int(year), int(month)
            self.period_manual = True
            self._set_detect("manual")
            self._invalidate()
            if self.prior_paths or self.douzone_paths:
                self._align_period()
            return self.state()

    def toggle_picker(self):
        with self.lock:
            self.picker_open = not self.picker_open
            return self.state()

    def set_account_codes(self, text):
        with self.lock:
            if text == self.account_codes:
                return self.state()
            self.account_codes = text
            had_result = self.run is not None
            self._invalidate()
            if self.prior_paths or self.douzone_paths:
                # keep the "inputs changed" message when a result was just dropped
                self._align_period(show_banner=not had_result)
            return self.state()

    def add_files(self, kind):
        """Open the file dialog, then classify the chosen files in the background."""
        with self.lock:
            if self.adding:
                self._banner("warning", "파일 분석 중", "현재 파일 분석이 끝난 뒤 다시 추가해주세요.")
                return self.state()
        paths = self.dialogs.open_files()
        return self.add_paths(kind, paths or [])

    def add_paths(self, kind, paths):
        paths = [str(p) for p in paths]
        with self.lock:
            if not paths or self.adding:
                return self.state()
            self.period_manual = False
            self.adding = True
            self.pending = [Path(p).name for p in paths]
            self._banner("running", "파일 분석 중", "Excel 구조와 회계월을 확인하고 있습니다. 큰 파일은 시간이 걸릴 수 있습니다.")
            codes = parse_account_codes(self.account_codes)
            prior, raw = tuple(self.prior_paths), tuple(self.douzone_paths)
        threading.Thread(target=self._classify, args=(kind, paths, codes, prior, raw), daemon=True).start()
        return self.state()

    def _classify(self, requested, paths, codes, prior, raw):
        classified, detected = [], {}
        try:
            for p in paths:
                try:
                    classified.append((p, classify_excel_input(p), None))
                except Exception as e:
                    classified.append((p, None, str(e)))
            for p in prior + tuple(p for p, k, e in classified if k == "prior" and not e and p not in prior):
                try: detected[p] = cached_periods(self.period_cache, "prior", p, codes)
                except Exception: detected[p] = []
            for p in raw + tuple(p for p, k, e in classified if k == "douzone" and not e and p not in raw):
                try: detected[p] = cached_periods(self.period_cache, "douzone", p, codes)
                except Exception: detected[p] = []
            error = None
        except Exception as e:
            error = str(e)
        with self.lock:
            try:
                if error:
                    self._banner("error", "파일 분석 실패", error)
                    self._notice("error", "파일 분석 실패", error)
                else:
                    self._finish_files(classified, requested, detected)
            finally:
                self.pending = []
                self.adding = False

    def _finish_files(self, classified, requested, detected):
        added, moved, rejected = False, [], []
        for p, kind, error in classified:
            name = Path(p).name
            if error:
                rejected.append(f"{name}: 파일을 읽지 못함 ({error})"); continue
            if kind == "prior":
                if p not in self.prior_paths:
                    self.prior_paths.append(p); added = True
                    if requested != "prior": moved.append(f"{name} → 명세서")
            elif kind == "douzone":
                if p not in self.douzone_paths:
                    self.douzone_paths.append(p); added = True
                    if requested != "douzone": moved.append(f"{name} → 더존 원장")
            elif kind == "ambiguous":
                rejected.append(f"{name}: 명세서와 더존 구조가 함께 보여 자동 분류하지 않음")
            else:
                rejected.append(f"{name}: 명세서/더존 원장 구조를 식별하지 못함")
        if added:
            self._invalidate()
            self._align_period(detected=detected)
        else:
            self._banner("idle", "대사 전", IDLE_HINT)
        if moved:
            self._banner("idle", "파일 자동 분류",
                         "잘못된 칸에서 선택한 파일을 구조에 맞게 자동 배치했습니다: " + " · ".join(moved))
        if rejected:
            self._notice("warning", "파일 자동 분류 확인",
                         "다음 파일은 안전하게 자동 분류할 수 없어 추가하지 않았습니다.\n\n" + "\n".join(rejected))

    def remove_file(self, kind, index):
        with self.lock:
            paths = self.prior_paths if kind == "prior" else self.douzone_paths
            if self.adding or not (0 <= int(index) < len(paths)):
                return self.state()
            del paths[int(index)]
            self._after_removed()
            return self.state()

    def clear_files(self, kind):
        with self.lock:
            paths = self.prior_paths if kind == "prior" else self.douzone_paths
            if self.adding or not paths:
                return self.state()
            paths.clear()
            self._after_removed()
            return self.state()

    def _after_removed(self):
        self.period_manual = False
        self._invalidate()
        self._align_period()

    def open_inputs(self):
        with self.lock:
            self.inputs_open = True
            return self.state()

    # -------------------------------------------------------------------- run
    def run_reconciliation(self):
        with self.lock:
            if self.adding:
                self._notice("info", "파일 분석 중", "파일 분석이 완료된 후 대사를 실행해주세요.")
                return self.state()
            if self.running:
                return self.state()
            if not self.prior_paths or not self.douzone_paths:
                self._notice("warning", "파일 필요", "명세서와 더존 원장을 각각 1개 이상 추가해주세요.")
                return self.state()
            self._invalidate()
            if self._align_period(show_banner=False) is False:
                self._banner("error", "대사 중단", "선택한 회계월과 파일의 월이 맞지 않습니다.")
                self._notice("error", "회계월 확인",
                             "전월 말 명세서와 당월 더존 원장이 이어지는 월인지, 선택한 회계월이 맞는지 확인해주세요.\n"
                             "파일의 월과 선택한 월을 확인해주세요.")
                return self.state()
            period = self._period()
            if period is None:
                self._banner("error", "대사 중단", "회계월을 확인한 뒤 다시 실행해주세요.")
                self._notice("error", "대사 중단", "회계월을 확인한 뒤 다시 실행해주세요.")
                return self.state()
            codes = parse_account_codes(self.account_codes)
            prior, raw = list(self.prior_paths), list(self.douzone_paths)
            self.running = True
            self._banner("running", "대사 중", "사전검사와 자동 대사를 진행하고 있습니다. 큰 파일은 시간이 걸릴 수 있습니다.")
        threading.Thread(target=self._do_run, args=(period, prior, raw, codes), daemon=True).start()
        return self.state()

    def _do_run(self, period, prior, raw, codes):
        sources = prior + raw
        try:
            before = snapshot_file_digests(sources)
            dup = find_duplicate_files(prior, raw, before)
            if dup:
                a, b = dup[0]
                raise ValueError(f"동일한 파일 내용이 중복 추가되었습니다: {a[1]} / {b[1]}")
            run = run_reconciliation(prior, raw, codes, period)
            after = snapshot_file_digests(sources)
            if before != after:
                raise ValueError("대사 실행 중 입력 Excel이 변경되었습니다. 파일 저장이 끝난 뒤 다시 실행해주세요.")
            with self.lock:
                self._finish_run(period, sources, run, after)
        except Exception as e:
            with self.lock:
                self.running = False
                self._banner("error", "대사 중단", "입력 내용을 확인한 뒤 다시 실행해주세요.")
                self._notice("error", "대사 중단", str(e))

    def _finish_run(self, period, sources, run, digests):
        self.results, self.new_items, self.issues = run.results, run.new_items, run.issues
        self.run, self.counts = run, run.counts
        rows = []
        for i, x in enumerate(self.results):
            matched = x.status in AUTO
            amount = int(x.closing_balance if x.status == Status.SIGNED_NET_AUTO else x.prior.amount)
            rows.append({"id": i, "vendor": x.prior.vendor_name, "amount": f"{amount:,}",
                         "status": x.status.value, "matched": matched, "reason": x.reason,
                         "source": x.prior.source.owner or x.prior.source.file_name})
        self.rows = rows
        self.preflight = ({"text": f"입력 확인 {len(self.issues):,}건 · 해당 행은 자동 대사에서 제외했습니다.", "ok": False}
                          if self.issues else
                          {"text": "사전검사 통과 · 자동 제외된 입력 이상이 없습니다.", "ok": True})
        self.last_period, self.last_source_paths, self.last_source_digests = period, sources, digests
        b = reconciliation_breakdown(self.counts, len(self.new_items), len(self.issues))
        detail = completion_detail(period, run.prior_count, self.counts, len(self.new_items), len(self.issues))
        if run.unseen_account_codes:
            detail += " · 이 달 전표가 없는 계정코드: " + ", ".join(run.unseen_account_codes)
        flagged = b["review"] or b["issues"] or run.unseen_account_codes
        self._banner("success", "대사 완료 · 확인 항목 있음" if flagged else "대사 완료", detail)
        self.inputs_open = False
        self.running = False

    # ----------------------------------------------------------------- export
    def _guard_export(self, need_run=True):
        if self.last_period is None or self.run is None:
            self._notice("warning", "대사 필요", "대사를 먼저 실행해주세요.")
            return False
        changed = changed_snapshot_paths(self.last_source_digests)
        if changed:
            self._invalidate()
            self._notice("warning", "입력 파일 변경",
                         "대사 후 원본 Excel이 변경되거나 사라졌습니다: " + ", ".join(changed) + "\n다시 대사를 실행해주세요.")
            return False
        return True

    def export_month_end(self):
        with self.lock:
            if not self._guard_export():
                return self.state()
            p = self.last_period
        path = self.dialogs.save_file(f"{p.year}_{p.month:02d}_당월명세서_검토용.xlsx")
        if not path:
            return self.state()
        with self.lock:
            try:
                run = self.run
                notices = write_month_end_statement(
                    path, self.prior_paths, run.results, run.new_items, run.issues, p,
                    standalone_debits=run.standalone_debits,
                    standalone_debit_sources=run.standalone_debit_sources,
                    journal_items=run.journal_items, journal_sources=run.journal_sources,
                    source_paths=self.last_source_paths)
                self._banner("success", "당월 명세서 초안 생성",
                             f"{Path(path).name} · 전월 양식 기반, 검토필요·변경내역 시트는 삭제 가능합니다.")
                if notices:
                    shown = "\n".join("· " + t for t in notices[:12])
                    more = f"\n· 외 {len(notices)-12}건 (검토필요 시트 참고)" if len(notices) > 12 else ""
                    self._notice("info", "당월 명세서 생성 완료 · 확인할 곳이 있습니다",
                                 "명세서는 만들었습니다. 읽을 수 없었거나 그대로 옮길 수 없었던 부분입니다. "
                                 "항목마다 어떻게 처리했는지 적혀 있습니다.\n\n" + shown + more)
            except PermissionError as e:
                self._notice("error", "파일 접근 거부",
                             "저장할 결과 파일 또는 관련 Excel 파일이 열려 있어 접근이 거부됐을 수 있습니다.\n"
                             "Excel에서 해당 파일을 닫은 뒤 다시 저장해 주세요.\n\n"
                             "계속 실패하면 다른 파일명이나 쓰기 가능한 폴더에 저장해 주세요.\n"
                             "파일·폴더의 읽기 전용 설정과 쓰기 권한도 확인해 주세요.\n\n"
                             f"문제 파일: {e.filename or path}")
                return self.state()
            except ValueError as e:
                self._banner("error", "당월 명세서 생성 불가", "전월 양식을 유지할 수 없어 생성하지 않았습니다.")
                self._notice("error", "당월 명세서 생성 실패",
                             f"{e}\n\n전월 양식을 그대로 유지할 수 없어 당월 명세서를 만들지 않았습니다. "
                             "다른 형태의 파일로 대신 만들지 않습니다.\n위 내용을 고친 뒤 다시 시도해 주세요. "
                             "대사 결과만 필요하면 '대사 결과 저장'을 사용하세요.")
                return self.state()
            except Exception as e:
                self._notice("error", "당월 명세서 생성 실패", str(e))
                return self.state()
        try:
            self.dialogs.open_folder(path)
        except OSError:
            with self.lock:
                self._notice("warning", "저장 완료 · 폴더 열기 실패",
                             "당월 명세서는 정상적으로 저장했습니다.\n저장 폴더를 자동으로 열지 못했습니다. "
                             f"아래 위치를 확인해 주세요.\n\n저장 위치: {Path(path).resolve()}")
        return self.state()

    def export_result(self):
        with self.lock:
            if not self._guard_export():
                return self.state()
            p = self.last_period
        path = self.dialogs.save_file(f"{p.year}_{p.month:02d}_명세서_대사결과.xlsx")
        if not path:
            return self.state()
        with self.lock:
            try:
                write_result(path, self.results, self.new_items, self.issues, self.last_source_paths, p.label)
                self._banner("success", "결과 저장 완료", f"{Path(path).name} · 원본 Excel은 수정하지 않았습니다.")
            except Exception as e:
                self._notice("error", "저장 실패", str(e))
            return self.state()

    # ----------------------------------------------------------------- update
    def check_update(self):
        """Look for a newer release (background-safe; never raises)."""
        if not updater.can_self_update():
            return self.state()
        release = updater.check_latest()
        with self.lock:
            if release is not None:
                self.update_release = release
        return self.state()

    def _update_view(self):
        rel = self.update_release
        return {"available": rel is not None, "date": rel.date if rel else "",
                "current": updater.version_label(), **self.update_state}

    def start_update(self):
        with self.lock:
            rel = self.update_release
            if rel is None or self.running or self.update_state["busy"]:
                return self.state()
            self.update_state = {"busy": True, "percent": 0, "error": ""}
        threading.Thread(target=self._do_update, args=(rel,), daemon=True).start()
        return self.state()

    def _do_update(self, release):
        import sys
        exe = Path(sys.executable)
        staged = updater.stage_path(exe)
        try:
            updater.stop_old_swap_scripts()
            fresh = updater.fetch_release(force=True)
            target = fresh if fresh is not None and updater.is_newer(updater.current_sha(), fresh.sha) else release

            def progress(done, total):
                with self.lock:
                    self.update_state["percent"] = int(done * 100 / total) if total else 0
            updater.download(target, staged, progress)
            updater.launch_finisher(exe, staged)
        except Exception as e:
            try:
                staged.unlink()
            except OSError:
                pass
            with self.lock:
                self.update_state = {"busy": False, "percent": 0,
                                     "error": str(e) or "내려받기에 실패했습니다. 인터넷 연결을 확인해 주세요."}
            return
        self.quit()
