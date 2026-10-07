from __future__ import annotations
import tkinter as tk
from collections import Counter
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from adapters.excel.reader import read_douzone, read_prior
from adapters.excel.writer import write_result
from domain.models import Status
from domain.period import AccountingPeriod
from domain.reconciliation import new_payables, reconcile

BG="#F5F5F7"; CARD="#FFFFFF"; TEXT="#1D1D1F"; MUTED="#6E6E73"; ACCENT="#0071E3"; WARN="#B45309"

class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.title("미지급금 대사"); self.geometry("980x720"); self.minsize(880,650); self.configure(bg=BG)
        today=date.today()
        self.year=tk.IntVar(value=today.year); self.month=tk.IntVar(value=today.month)
        self.prior_paths=[]; self.douzone_path=tk.StringVar(); self.account_codes=tk.StringVar(value="25301")
        self.period_text=tk.StringVar(); self.status_text=tk.StringVar(value="대상 회계월과 파일을 확인한 뒤 대사를 시작하세요.")
        self.results=[]; self.new_items=[]; self.issues=[]; self._build(); self._update_period()

    def _build(self):
        s=ttk.Style(self); s.theme_use("clam")
        s.configure("TButton",font=("Segoe UI",10),padding=(14,9)); s.configure("TEntry",padding=8)
        s.configure("Accent.TButton",font=("Segoe UI Semibold",10),padding=(18,10),foreground="white",background=ACCENT)
        root=tk.Frame(self,bg=BG); root.pack(fill="both",expand=True,padx=44,pady=30)
        tk.Label(root,text="미지급금 대사",font=("Segoe UI Semibold",26),bg=BG,fg=TEXT).pack(anchor="w")
        tk.Label(root,text="정상 건은 숨기고 사람이 확인해야 할 항목만 선명하게 보여줍니다.",font=("Segoe UI",11),bg=BG,fg=MUTED).pack(anchor="w",pady=(4,18))

        period=tk.Frame(root,bg=CARD,highlightthickness=1,highlightbackground="#E5E5EA"); period.pack(fill="x",pady=(0,12))
        left=tk.Frame(period,bg=CARD); left.pack(side="left",padx=22,pady=16)
        tk.Label(left,text="대상 회계월",font=("Segoe UI Semibold",10),bg=CARD,fg=TEXT).pack(anchor="w")
        controls=tk.Frame(left,bg=CARD); controls.pack(anchor="w",pady=(7,0))
        ttk.Spinbox(controls,from_=2020,to=2100,textvariable=self.year,width=7,command=self._update_period).pack(side="left")
        tk.Label(controls,text="년",bg=CARD,fg=TEXT).pack(side="left",padx=(4,10))
        ttk.Spinbox(controls,from_=1,to=12,textvariable=self.month,width=4,command=self._update_period).pack(side="left")
        tk.Label(controls,text="월",bg=CARD,fg=TEXT).pack(side="left",padx=4)
        self.year.trace_add("write",lambda *_:self._update_period()); self.month.trace_add("write",lambda *_:self._update_period())
        tk.Label(period,textvariable=self.period_text,font=("Segoe UI Semibold",12),bg=CARD,fg=ACCENT,justify="left").pack(side="left",padx=30)

        card=tk.Frame(root,bg=CARD,highlightthickness=1,highlightbackground="#E5E5EA"); card.pack(fill="x")
        tk.Label(card,text="전월 담당자 명세서",font=("Segoe UI Semibold",10),bg=CARD,fg=TEXT).grid(row=0,column=0,sticky="nw",padx=(22,12),pady=16)
        self.prior_list=tk.Listbox(card,height=4,font=("Segoe UI",9),selectmode="extended"); self.prior_list.grid(row=0,column=1,sticky="ew",pady=16)
        pf=tk.Frame(card,bg=CARD); pf.grid(row=0,column=2,padx=18,pady=16,sticky="n")
        ttk.Button(pf,text="파일 추가",command=self.pick_priors).pack(fill="x"); ttk.Button(pf,text="선택 제거",command=self.remove_priors).pack(fill="x",pady=(6,0))
        self._file_row(card,"더존 Raw (1개월~1년 등 임의 기간)",self.douzone_path,1)
        opt=tk.Frame(card,bg=CARD); opt.grid(row=2,column=0,columnspan=3,sticky="ew",padx=22,pady=(2,18))
        tk.Label(opt,text="미지급금 계정코드",font=("Segoe UI Semibold",10),bg=CARD,fg=TEXT).pack(side="left")
        ttk.Entry(opt,textvariable=self.account_codes,width=18).pack(side="left",padx=12)
        tk.Label(opt,text="여러 개면 쉼표로 구분",font=("Segoe UI",9),bg=CARD,fg=MUTED).pack(side="left")

        actions=tk.Frame(root,bg=BG); actions.pack(fill="x",pady=16)
        ttk.Button(actions,text="사전검사 + 대사 시작",style="Accent.TButton",command=self.run).pack(side="left")
        self.export_btn=ttk.Button(actions,text="결과 Excel 저장",command=self.export,state="disabled"); self.export_btn.pack(side="left",padx=10)
        self.summary=tk.Frame(root,bg=BG); self.summary.pack(fill="x",pady=(2,8)); self._summary({})
        self.preflight=tk.Label(root,text="",font=("Segoe UI",10),bg=BG,fg=WARN,justify="left",anchor="w"); self.preflight.pack(fill="x",pady=(8,0))
        tk.Label(root,textvariable=self.status_text,font=("Segoe UI",10),bg=BG,fg=MUTED,anchor="w").pack(fill="x",pady=(8,0))
        self.detail=ttk.Treeview(root,columns=("owner","vendor","amount","status","reason"),show="headings",height=7)
        for col,title,width in (("owner","담당자",110),("vendor","거래처",150),("amount","금액",110),("status","상태",120),("reason","사유",320)):
            self.detail.heading(col,text=title); self.detail.column(col,width=width,anchor="w")
        self.detail.pack(fill="both",expand=True,pady=(10,0))
        tk.Label(root,text="안전 모드  ·  더존 직접 조작 없음  ·  날짜는 대사키로 사용하지 않음  ·  원본 덮어쓰기 금지",font=("Segoe UI",9),bg=BG,fg=MUTED).pack(anchor="w",pady=(18,0))

    def _update_period(self):
        try:
            p=AccountingPeriod(int(self.year.get()),int(self.month.get())); prev=p.previous()
            self.period_text.set(f"{prev.label} 명세  →  {p.label} 전표\n※ Raw에서는 이 회계월만 추출하고, 개별 대사키에는 날짜를 쓰지 않습니다.")
        except Exception: self.period_text.set("올바른 연/월을 선택해주세요.")

    def _file_row(self,parent,label,var,row):
        parent.grid_columnconfigure(1,weight=1)
        tk.Label(parent,text=label,font=("Segoe UI Semibold",10),bg=CARD,fg=TEXT).grid(row=row,column=0,sticky="w",padx=(22,12),pady=16)
        ttk.Entry(parent,textvariable=var).grid(row=row,column=1,sticky="ew",pady=16)
        ttk.Button(parent,text="파일 선택",command=lambda:self.pick(var)).grid(row=row,column=2,padx=18,pady=16)

    def pick_priors(self):
        paths=filedialog.askopenfilenames(filetypes=[("Excel","*.xlsx *.xlsm")])
        for p in paths:
            if p not in self.prior_paths:
                self.prior_paths.append(p); self.prior_list.insert("end",Path(p).name)

    def remove_priors(self):
        for i in reversed(self.prior_list.curselection()):
            self.prior_list.delete(i); self.prior_paths.pop(i)

    def pick(self,var):
        p=filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xlsm")])
        if p: var.set(p)

    def _summary(self,counts):
        for w in self.summary.winfo_children(): w.destroy()
        cards=[("자동 대사",counts.get(Status.MATCHED,0)),("대사 예외",sum(v for k,v in counts.items() if k!=Status.MATCHED)),("입력 확인",len(self.issues)),("신규 미지급",len(self.new_items))]
        for i,(name,value) in enumerate(cards):
            box=tk.Frame(self.summary,bg=CARD,highlightthickness=1,highlightbackground="#E5E5EA"); box.grid(row=0,column=i,sticky="nsew",padx=(0 if i==0 else 7,0)); self.summary.grid_columnconfigure(i,weight=1)
            tk.Label(box,text=name,font=("Segoe UI",10),bg=CARD,fg=MUTED).pack(anchor="w",padx=16,pady=(13,2))
            tk.Label(box,text=f"{value:,}",font=("Segoe UI Semibold",22),bg=CARD,fg=TEXT).pack(anchor="w",padx=16,pady=(0,13))

    def run(self):
        if not self.prior_paths or not self.douzone_path.get(): messagebox.showwarning("파일 필요","전월 담당자 명세서를 1개 이상 추가하고 더존 Raw를 선택해주세요."); return
        try:
            p=AccountingPeriod(int(self.year.get()),int(self.month.get()))
            codes={x.strip() for x in self.account_codes.get().split(",") if x.strip()}
            prior_items=[]; prior_issues=[]
            for path in self.prior_paths:
                rr=read_prior(path,Path(path).stem); prior_items.extend(rr.items); prior_issues.extend(rr.issues)
            dz=read_douzone(self.douzone_path.get(),codes or None,p)
            self.issues=prior_issues+dz.issues; self.results=reconcile(prior_items,dz.items); self.new_items=new_payables(dz.items)
            counts=Counter(x.status for x in self.results); self._summary(counts)
            exc=sum(v for k,v in counts.items() if k!=Status.MATCHED)
            self.preflight.config(text=(f"⚠ 입력 형식 확인 {len(self.issues):,}건 — 해당 행은 자동대사에서 제외했습니다." if self.issues else "✓ 사전검사 통과 — 의심스러운 금액 형식 없음"),fg=(WARN if self.issues else "#2E7D32"))
            for row in self.detail.get_children(): self.detail.delete(row)
            for x in self.results:
                if x.status != Status.MATCHED:
                    self.detail.insert("", "end", values=(x.prior.source.owner or x.prior.source.file_name, x.prior.vendor_name, f"{int(x.prior.amount):,}", x.status.value, x.reason))
            self.status_text.set(f"{p.label} 대사 완료 · 전월 {len(prior_items):,}건 · 대사 예외 {exc:,}건 · 신규 {len(self.new_items):,}건")
            self.export_btn.config(state="normal")
        except Exception as e: messagebox.showerror("대사 중단",str(e))

    def export(self):
        if not self.results and not self.issues: return
        p=AccountingPeriod(int(self.year.get()),int(self.month.get()))
        path=filedialog.asksaveasfilename(defaultextension=".xlsx",initialfile=f"{p.year}_{p.month:02d}_미지급금_대사결과.xlsx",filetypes=[("Excel","*.xlsx")])
        if not path:return
        try:
            write_result(path,self.results,self.new_items,self.issues,self.prior_paths+[self.douzone_path.get()],p.label)
            self.status_text.set(f"결과 저장 완료 · {Path(path).name}"); messagebox.showinfo("저장 완료","원본은 수정하지 않았습니다.\n확인필요/입력데이터확인 시트를 먼저 확인해주세요.")
        except Exception as e: messagebox.showerror("저장 실패",str(e))

if __name__=="__main__": App().mainloop()
