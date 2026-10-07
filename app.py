from __future__ import annotations

import tkinter as tk
from collections import Counter
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from excel_io import read_douzone, read_prior, write_result
from reconciliation import new_payables, reconcile
from models import Status


BG = "#F5F5F7"
CARD = "#FFFFFF"
TEXT = "#1D1D1F"
MUTED = "#6E6E73"
ACCENT = "#0071E3"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("미지급금 대사")
        self.geometry("920x650")
        self.minsize(820, 590)
        self.configure(bg=BG)
        self.prior_path = tk.StringVar()
        self.douzone_path = tk.StringVar()
        self.account_codes = tk.StringVar(value="25301")
        self.status_text = tk.StringVar(value="파일 두 개를 선택한 뒤 대사를 시작하세요.")
        self.results = []
        self.new_items = []
        self._build()

    def _build(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TButton", font=("Segoe UI", 10), padding=(14, 9))
        style.configure("Accent.TButton", font=("Segoe UI Semibold", 10), padding=(18, 10),
                        foreground="white", background=ACCENT)
        style.map("Accent.TButton", background=[("active", "#0077ED")])
        style.configure("TEntry", padding=8)

        root = tk.Frame(self, bg=BG)
        root.pack(fill="both", expand=True, padx=44, pady=34)
        tk.Label(root, text="미지급금 대사", font=("Segoe UI Semibold", 26), bg=BG, fg=TEXT).pack(anchor="w")
        tk.Label(root, text="정상 건은 자동으로 정리하고, 확인이 필요한 예외만 보여줍니다.",
                 font=("Segoe UI", 11), bg=BG, fg=MUTED).pack(anchor="w", pady=(4, 22))

        card = tk.Frame(root, bg=CARD, highlightthickness=1, highlightbackground="#E5E5EA")
        card.pack(fill="x")
        self._file_row(card, "전월 미지급 세부명세", self.prior_path, 0)
        self._file_row(card, "당월 더존 전표출력", self.douzone_path, 1)

        opt = tk.Frame(card, bg=CARD)
        opt.grid(row=2, column=0, columnspan=3, sticky="ew", padx=22, pady=(4, 20))
        tk.Label(opt, text="미지급금 계정코드", font=("Segoe UI Semibold", 10), bg=CARD, fg=TEXT).pack(side="left")
        ttk.Entry(opt, textvariable=self.account_codes, width=18).pack(side="left", padx=12)
        tk.Label(opt, text="여러 개면 쉼표로 구분 · 비우면 계정명에 '미지급' 포함 항목",
                 font=("Segoe UI", 9), bg=CARD, fg=MUTED).pack(side="left")

        actions = tk.Frame(root, bg=BG)
        actions.pack(fill="x", pady=18)
        ttk.Button(actions, text="대사 시작", style="Accent.TButton", command=self.run_reconcile).pack(side="left")
        self.export_btn = ttk.Button(actions, text="결과 Excel 저장", command=self.export, state="disabled")
        self.export_btn.pack(side="left", padx=10)

        self.summary = tk.Frame(root, bg=BG)
        self.summary.pack(fill="x", pady=(6, 10))
        self._render_summary({})

        tk.Label(root, textvariable=self.status_text, font=("Segoe UI", 10), bg=BG, fg=MUTED,
                 anchor="w").pack(fill="x", pady=(8, 0))
        tk.Label(root, text="안전 모드 · 더존 직접 조작 없음 · 원본 Excel 덮어쓰기 금지",
                 font=("Segoe UI", 9), bg=BG, fg=MUTED).pack(anchor="w", pady=(18, 0))

    def _file_row(self, parent, label, var, row):
        parent.grid_columnconfigure(1, weight=1)
        tk.Label(parent, text=label, font=("Segoe UI Semibold", 10), bg=CARD, fg=TEXT).grid(
            row=row, column=0, sticky="w", padx=(22, 12), pady=18)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", pady=18)
        ttk.Button(parent, text="파일 선택", command=lambda: self.pick(var)).grid(
            row=row, column=2, padx=18, pady=18)

    def pick(self, var):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if path:
            var.set(path)

    def _render_summary(self, counts):
        for w in self.summary.winfo_children():
            w.destroy()
        cards = [
            ("자동 대사", counts.get(Status.MATCHED, 0)),
            ("확인 필요", sum(v for k, v in counts.items() if k != Status.MATCHED)),
            ("신규 미지급", len(self.new_items)),
        ]
        for i, (name, value) in enumerate(cards):
            box = tk.Frame(self.summary, bg=CARD, highlightthickness=1, highlightbackground="#E5E5EA")
            box.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 7, 0))
            self.summary.grid_columnconfigure(i, weight=1)
            tk.Label(box, text=name, font=("Segoe UI", 10), bg=CARD, fg=MUTED).pack(anchor="w", padx=18, pady=(15, 2))
            tk.Label(box, text=str(value), font=("Segoe UI Semibold", 23), bg=CARD, fg=TEXT).pack(anchor="w", padx=18, pady=(0, 15))

    def run_reconcile(self):
        if not self.prior_path.get() or not self.douzone_path.get():
            messagebox.showwarning("파일 필요", "전월 명세서와 더존 전표출력 파일을 모두 선택해주세요.")
            return
        try:
            codes = {x.strip() for x in self.account_codes.get().split(",") if x.strip()}
            prior = read_prior(self.prior_path.get())
            journal = read_douzone(self.douzone_path.get(), codes or None)
            self.results = reconcile(prior, journal)
            self.new_items = new_payables(journal)
            counts = Counter(x.status for x in self.results)
            self._render_summary(counts)
            exceptions = sum(v for k, v in counts.items() if k != Status.MATCHED)
            self.status_text.set(f"대사 완료 · 전월 {len(prior):,}건 중 확인 필요 {exceptions:,}건")
            self.export_btn.config(state="normal")
        except Exception as e:
            messagebox.showerror("대사 실패", str(e))

    def export(self):
        if not self.results:
            return
        default = "미지급금_대사결과.xlsx"
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=default,
                                            filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            write_result(path, self.results, self.new_items, [self.prior_path.get(), self.douzone_path.get()])
            self.status_text.set(f"결과 저장 완료 · {Path(path).name}")
            messagebox.showinfo("저장 완료", "원본은 수정하지 않았습니다.\n결과 파일을 별도로 생성했습니다.")
        except Exception as e:
            messagebox.showerror("저장 실패", str(e))


if __name__ == "__main__":
    App().mainloop()
