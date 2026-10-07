from __future__ import annotations
import tkinter as tk
import hashlib
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from adapters.excel.writer import write_result
from domain.models import Status
from domain.period import AccountingPeriod
from application.service import run_reconciliation

BG="#F5F5F7"; CARD="#FFFFFF"; TEXT="#1D1D1F"; MUTED="#6E6E73"; ACCENT="#007AFF"; WARN="#B45309"; BORDER="#D2D2D7"; SOFT="#E8E8ED"

def file_digest(path):
    h=hashlib.sha256()
    with open(path,"rb") as f:
        for chunk in iter(lambda:f.read(1024*1024),b""):
            h.update(chunk)
    return h.hexdigest()

def snapshot_file_digests(paths):
    return {str(Path(path).resolve()):file_digest(path) for path in paths}

def changed_snapshot_paths(snapshot):
    changed=[]
    for path,digest in snapshot.items():
        p=Path(path)
        if not p.exists() or file_digest(p)!=digest:
            changed.append(p.name)
    return changed

def find_duplicate_files(prior_paths, douzone_paths, digests=None):
    seen={}; duplicates=[]
    digests=digests or snapshot_file_digests(list(prior_paths)+list(douzone_paths))
    for kind, paths in (("전월 명세",prior_paths),("더존 Raw",douzone_paths)):
        for path in paths:
            digest=digests[str(Path(path).resolve())]
            if digest in seen:
                duplicates.append((seen[digest],(kind,Path(path).name)))
            else:
                seen[digest]=(kind,Path(path).name)
    return duplicates

def statement_confirmation_text(period):
    return f"선택한 전월 명세서가 모두 {period.previous().label} 마감본임을 확인했습니다."

def validate_statement_confirmation(confirmed, period):
    if not confirmed:
        raise ValueError(
            f"전월 명세서 확인이 필요합니다. 선택한 파일이 모두 {period.previous().label} 마감본인지 확인해주세요."
        )

class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.title("명세서 대사"); self.geometry("980x720"); self.minsize(880,650); self.configure(bg=BG)
        today=date.today()
        self.year=tk.IntVar(value=today.year); self.month=tk.IntVar(value=today.month)
        self.prior_paths=[]; self.douzone_paths=[]; self.account_codes=tk.StringVar(value="25301")
        self.period_text=tk.StringVar(); self.statement_confirm_text=tk.StringVar(); self.statement_confirmed=tk.BooleanVar(value=False)
        self.status_text=tk.StringVar(value="대상 회계월과 파일을 확인한 뒤 대사를 시작하세요.")
        self.results=[]; self.new_items=[]; self.issues=[]; self.last_period=None; self.last_source_paths=[]; self.last_source_digests={}; self._build(); self._update_period()

    def _build(self):
        s=ttk.Style(self); s.theme_use("clam")
        s.configure("TButton",font=("Segoe UI",10),padding=(14,9),background="#FFFFFF",bordercolor=BORDER); s.map("TButton",background=[("active",SOFT)]); s.configure("TEntry",padding=8,fieldbackground="#FFFFFF",bordercolor=BORDER)
        s.configure("Accent.TButton",font=("Segoe UI Semibold",10),padding=(18,10),foreground="white",background=ACCENT,bordercolor=ACCENT); s.map("Accent.TButton",background=[("active","#0066CC")])
        root=tk.Frame(self,bg=BG); root.pack(fill="both",expand=True,padx=44,pady=30)
        tk.Label(root,text="명세서 대사",font=("Segoe UI Semibold",28),bg=BG,fg=TEXT).pack(anchor="w")
        tk.Label(root,text="전월 명세서와 더존 전표를 대사해 차월 명세서 초안을 만들고, 확인이 필요한 항목만 보여줍니다.",font=("Segoe UI",11),bg=BG,fg=MUTED).pack(anchor="w",pady=(4,18))

        period=tk.Frame(root,bg=CARD,highlightthickness=1,highlightbackground=BORDER); period.pack(fill="x",pady=(0,12))
        left=tk.Frame(period,bg=CARD); left.pack(side="left",padx=22,pady=16)
        tk.Label(left,text="대상 회계월",font=("Segoe UI Semibold",10),bg=CARD,fg=TEXT).pack(anchor="w")
        controls=tk.Frame(left,bg=CARD); controls.pack(anchor="w",pady=(7,0))
        ttk.Spinbox(controls,from_=2020,to=2100,textvariable=self.year,width=7,command=self._update_period).pack(side="left")
        tk.Label(controls,text="년",bg=CARD,fg=TEXT).pack(side="left",padx=(4,10))
        ttk.Spinbox(controls,from_=1,to=12,textvariable=self.month,width=4,command=self._update_period).pack(side="left")
        tk.Label(controls,text="월",bg=CARD,fg=TEXT).pack(side="left",padx=4)
        self.year.trace_add("write",lambda *_:self._on_period_changed()); self.month.trace_add("write",lambda *_:self._on_period_changed())
        tk.Label(period,textvariable=self.period_text,font=("Segoe UI Semibold",12),bg=CARD,fg=ACCENT,justify="left").pack(side="left",padx=30)

        card=tk.Frame(root,bg=CARD,highlightthickness=1,highlightbackground="#E5E5EA"); card.pack(fill="x")
        tk.Label(card,text="전월 명세서",font=("Segoe UI Semibold",10),bg=CARD,fg=TEXT).grid(row=0,column=0,sticky="nw",padx=(22,12),pady=16)
        self.prior_list=tk.Listbox(card,height=4,font=("Segoe UI",9),selectmode="extended"); self.prior_list.grid(row=0,column=1,sticky="ew",pady=16)
        pf=tk.Frame(card,bg=CARD); pf.grid(row=0,column=2,padx=18,pady=16,sticky="n")
        ttk.Button(pf,text="파일 추가",command=self.pick_priors).pack(fill="x"); ttk.Button(pf,text="선택 제거",command=self.remove_priors).pack(fill="x",pady=(6,0))
        tk.Label(card,text="더존 Raw (여러 파일 가능)",font=("Segoe UI Semibold",10),bg=CARD,fg=TEXT).grid(row=1,column=0,sticky="nw",padx=(22,12),pady=16)
        self.douzone_list=tk.Listbox(card,height=4,font=("Segoe UI",9),selectmode="extended"); self.douzone_list.grid(row=1,column=1,sticky="ew",pady=16)
        df=tk.Frame(card,bg=CARD); df.grid(row=1,column=2,padx=18,pady=16,sticky="n")
        ttk.Button(df,text="파일 추가",command=self.pick_douzone).pack(fill="x"); ttk.Button(df,text="선택 제거",command=self.remove_douzone).pack(fill="x",pady=(6,0))
        opt=tk.Frame(card,bg=CARD); opt.grid(row=2,column=0,columnspan=3,sticky="ew",padx=22,pady=(2,8))
        tk.Label(opt,text="미지급금 계정코드",font=("Segoe UI Semibold",10),bg=CARD,fg=TEXT).pack(side="left")
        ttk.Entry(opt,textvariable=self.account_codes,width=18).pack(side="left",padx=12)
        tk.Label(opt,text="여러 개면 쉼표로 구분",font=("Segoe UI",9),bg=CARD,fg=MUTED).pack(side="left")
        confirm=tk.Frame(card,bg=CARD); confirm.grid(row=3,column=0,columnspan=3,sticky="ew",padx=22,pady=(0,16))
        ttk.Checkbutton(confirm,textvariable=self.statement_confirm_text,variable=self.statement_confirmed).pack(anchor="w")
        tk.Label(confirm,text="※ 명세서 행의 날짜는 장기이월 때문에 월 검증에 사용하지 않습니다.",font=("Segoe UI",9),bg=CARD,fg=MUTED).pack(anchor="w",padx=(22,0),pady=(2,0))

        actions=tk.Frame(root,bg=BG); actions.pack(fill="x",pady=16)
        ttk.Button(actions,text="사전검사 + 대사 시작",style="Accent.TButton",command=self.run).pack(side="left")
        self.export_btn=ttk.Button(actions,text="결과 Excel 저장",command=self.export,state="disabled"); self.export_btn.pack(side="left",padx=10)
        self.summary=tk.Frame(root,bg=BG); self.summary.pack(fill="x",pady=(2,8)); self._summary({})
        self.preflight=tk.Label(root,text="",font=("Segoe UI",10),bg=BG,fg=WARN,justify="left",anchor="w"); self.preflight.pack(fill="x",pady=(8,0))
        tk.Label(root,textvariable=self.status_text,font=("Segoe UI",10),bg=BG,fg=MUTED,anchor="w").pack(fill="x",pady=(8,0))
        self.detail=ttk.Treeview(root,columns=("owner","vendor","amount","status","reason"),show="headings",height=7)
        for col,title,width in (("owner","원본 명세",150),("vendor","거래처",150),("amount","금액",110),("status","상태",130),("reason","사유",300)):
            self.detail.heading(col,text=title); self.detail.column(col,width=width,anchor="w")
        self.detail.pack(fill="both",expand=True,pady=(10,0))
        tk.Label(root,text="안전 모드  ·  더존 직접 조작 없음  ·  날짜는 대사키로 사용하지 않음  ·  원본 덮어쓰기 금지",font=("Segoe UI",9),bg=BG,fg=MUTED).pack(anchor="w",pady=(18,0))
        self.account_codes.trace_add("write",lambda *_:self._invalidate_results())

    def _update_period(self):
        try:
            p=AccountingPeriod(int(self.year.get()),int(self.month.get()))
            self.period_text.set(f"{p.previous().label} 명세서  →  {p.label} 더존 전표\n※ Raw는 선택 회계월로 필터링하고, 명세서 행 날짜는 대사키로 쓰지 않습니다.")
            self.statement_confirm_text.set(statement_confirmation_text(p))
        except Exception:
            self.period_text.set("올바른 연/월을 선택해주세요.")
            self.statement_confirm_text.set("전월 명세서의 마감월을 확인해주세요.")

    def _on_period_changed(self):
        self._update_period()
        self._invalidate_results(reset_statement_confirmation=True)

    def _invalidate_results(self, reset_statement_confirmation=False):
        if not hasattr(self,"export_btn"):
            return
        had_result=bool(self.results or self.new_items or self.issues or self.last_period)
        self.results=[]; self.new_items=[]; self.issues=[]; self.last_period=None; self.last_source_paths=[]; self.last_source_digests={}
        if reset_statement_confirmation:
            self.statement_confirmed.set(False)
        self.export_btn.config(state="disabled")
        self._summary({})
        self.preflight.config(text="")
        for row in self.detail.get_children():
            self.detail.delete(row)
        if had_result:
            self.status_text.set("입력 조건이 변경되었습니다. 다시 대사를 실행하세요.")

    def _file_row(self,parent,label,var,row):
        parent.grid_columnconfigure(1,weight=1)
        tk.Label(parent,text=label,font=("Segoe UI Semibold",10),bg=CARD,fg=TEXT).grid(row=row,column=0,sticky="w",padx=(22,12),pady=16)
        ttk.Entry(parent,textvariable=var).grid(row=row,column=1,sticky="ew",pady=16)
        ttk.Button(parent,text="파일 선택",command=lambda:self.pick(var)).grid(row=row,column=2,padx=18,pady=16)

    def pick_priors(self):
        paths=filedialog.askopenfilenames(filetypes=[("Excel","*.xlsx *.xlsm")])
        changed=False
        for p in paths:
            if p not in self.prior_paths:
                self.prior_paths.append(p); self.prior_list.insert("end",Path(p).name); changed=True
        if changed: self._invalidate_results(reset_statement_confirmation=True)

    def pick_douzone(self):
        paths=filedialog.askopenfilenames(filetypes=[("Excel","*.xlsx *.xlsm")])
        changed=False
        for p in paths:
            if p not in self.douzone_paths:
                self.douzone_paths.append(p); self.douzone_list.insert("end",Path(p).name); changed=True
        if changed: self._invalidate_results()

    def remove_douzone(self):
        selected=list(self.douzone_list.curselection())
        for i in reversed(selected):
            self.douzone_list.delete(i); self.douzone_paths.pop(i)
        if selected: self._invalidate_results()

    def remove_priors(self):
        selected=list(self.prior_list.curselection())
        for i in reversed(selected):
            self.prior_list.delete(i); self.prior_paths.pop(i)
        if selected: self._invalidate_results(reset_statement_confirmation=True)

    def pick(self,var):
        p=filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xlsm")])
        if p: var.set(p)

    def _summary(self,counts):
        for w in self.summary.winfo_children(): w.destroy()
        cards=[("자동 대사 완료",counts.get(Status.MATCHED,0)),("검토 필요",sum(v for k,v in counts.items() if k!=Status.MATCHED)+len(self.new_items)),("입력 확인",len(self.issues)),("당월 신규 명세(검토)",len(self.new_items))]
        for i,(name,value) in enumerate(cards):
            box=tk.Frame(self.summary,bg=CARD,highlightthickness=1,highlightbackground="#E5E5EA"); box.grid(row=0,column=i,sticky="nsew",padx=(0 if i==0 else 7,0)); self.summary.grid_columnconfigure(i,weight=1)
            tk.Label(box,text=name,font=("Segoe UI",10),bg=CARD,fg=MUTED).pack(anchor="w",padx=16,pady=(13,2))
            tk.Label(box,text=f"{value:,}",font=("Segoe UI Semibold",22),bg=CARD,fg=TEXT).pack(anchor="w",padx=16,pady=(0,13))

    def _duplicate_files(self):
        return find_duplicate_files(self.prior_paths,self.douzone_paths)

    def run(self):
        if not self.prior_paths or not self.douzone_paths:
            messagebox.showwarning("파일 필요","전월 명세서와 더존 Raw를 각각 1개 이상 추가해주세요."); return
        self._invalidate_results()
        self.status_text.set("사전검사 및 대사 실행 중...")
        self.update_idletasks()
        try:
            source_paths=list(self.prior_paths)+list(self.douzone_paths)
            before_digests=snapshot_file_digests(source_paths)
            dup=find_duplicate_files(self.prior_paths,self.douzone_paths,before_digests)
            if dup:
                a,b=dup[0]
                raise ValueError(f"동일한 파일 내용이 중복 추가되었습니다: {a[1]} / {b[1]}")
            p=AccountingPeriod(int(self.year.get()),int(self.month.get()))
            validate_statement_confirmation(self.statement_confirmed.get(),p)
            codes={x.strip() for x in self.account_codes.get().split(",") if x.strip()}
            run=run_reconciliation(list(self.prior_paths),list(self.douzone_paths),codes,p)
            after_digests=snapshot_file_digests(source_paths)
            if before_digests != after_digests:
                raise ValueError("대사 실행 중 입력 Excel이 변경되었습니다. 파일 저장이 끝난 뒤 다시 실행해주세요.")
            self.issues=run.issues; self.results=run.results; self.new_items=run.new_items
            counts=run.counts; self._summary(counts)
            exc=sum(v for k,v in counts.items() if k!=Status.MATCHED)
            self.preflight.config(text=(f"⚠ 입력 형식 확인 {len(self.issues):,}건 — 해당 행은 자동대사에서 제외했습니다." if self.issues else "✓ 사전검사 통과 — 자동 제외할 입력 이상 없음"),fg=(WARN if self.issues else "#2E7D32"))
            for row in self.detail.get_children(): self.detail.delete(row)
            for x in self.results:
                if x.status != Status.MATCHED:
                    self.detail.insert("", "end", values=(x.prior.source.owner or x.prior.source.file_name, x.prior.vendor_name, f"{int(x.prior.amount):,}", x.status.value, x.reason))
            self.last_period=p
            self.last_source_paths=source_paths
            self.last_source_digests=after_digests
            self.status_text.set(f"{p.label} 대사 완료 · 전월 명세 {run.prior_count:,}건 · 전월 검토 {exc:,}건 · 당월 신규 검토 {len(self.new_items):,}건")
            self.export_btn.config(state="normal")
        except Exception as e:
            self.status_text.set("대사가 완료되지 않았습니다. 입력 내용을 확인해주세요.")
            messagebox.showerror("대사 중단",str(e))

    def export(self):
        if self.last_period is None:
            messagebox.showwarning("대사 필요","현재 입력 조건으로 대사를 먼저 실행해주세요."); return
        changed=changed_snapshot_paths(self.last_source_digests)
        if changed:
            self._invalidate_results()
            messagebox.showwarning("입력 파일 변경","대사 후 입력 Excel이 변경되거나 사라졌습니다: "+", ".join(changed)+"\n다시 대사를 실행해주세요.")
            return
        p=self.last_period
        path=filedialog.asksaveasfilename(defaultextension=".xlsx",initialfile=f"{p.year}_{p.month:02d}_명세서_대사결과.xlsx",filetypes=[("Excel","*.xlsx")])
        if not path:return
        try:
            write_result(path,self.results,self.new_items,self.issues,self.last_source_paths,p.label)
            self.status_text.set(f"결과 저장 완료 · {Path(path).name}"); messagebox.showinfo("저장 완료","원본은 수정하지 않았습니다.\n확인필요/입력데이터확인 시트를 먼저 확인해주세요.")
        except Exception as e: messagebox.showerror("저장 실패",str(e))

if __name__=="__main__": App().mainloop()
