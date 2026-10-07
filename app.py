from __future__ import annotations
import tkinter as tk
import customtkinter as ctk
import hashlib
from collections import Counter
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from adapters.excel.reader import classify_excel_input, detect_douzone_periods, detect_statement_periods, read_douzone, read_prior
from adapters.excel.writer import write_result
from domain.models import Status
from domain.period import AccountingPeriod
from application.service import run_reconciliation

BG="#F5F5F7"
CARD="#FFFFFF"
TEXT="#1D1D1F"
MUTED="#6E6E73"
ACCENT="#007AFF"
ACCENT_HOVER="#0066CC"
GREEN="#1F7A4D"
GREEN_BG="#EFF8F3"
WARN="#A15C00"
WARN_BG="#FFF7E8"
RED="#B42318"
RED_BG="#FFF1F0"
BORDER="#E5E5EA"
BORDER_STRONG="#D2D2D7"
SOFT="#F2F2F7"
LIST_BG="#FAFAFC"
SELECT_BG="#DDEBFF"

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


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
    for kind, paths in (("명세서",prior_paths),("더존 Raw",douzone_paths)):
        for path in paths:
            digest=digests[str(Path(path).resolve())]
            if digest in seen:
                duplicates.append((seen[digest],(kind,Path(path).name)))
            else:
                seen[digest]=(kind,Path(path).name)
    return duplicates


def reconciliation_breakdown(counts, new_count=0, issue_count=0):
    matched=counts.get(Status.MATCHED,0)
    review=sum(v for k,v in counts.items() if k!=Status.MATCHED)+new_count
    return {
        "matched":matched,
        "review":review,
        "issues":issue_count,
        "new":new_count,
    }


def completion_detail(period, prior_count, counts, new_count=0, issue_count=0):
    b=reconciliation_breakdown(counts,new_count,issue_count)
    return (
        f"{period.label} · 명세서 {prior_count:,}건 처리 · "
        f"자동 대사 {b['matched']:,}건 · 검토 필요 {b['review']:,}건 · 입력 확인 {b['issues']:,}건"
    )


def suggest_reconciliation_period(current_period, statement_period_groups):
    """Infer the target accounting month directly from the latest statement month."""
    groups=[set(group) for group in statement_period_groups if group]
    if not groups:
        return None
    common=set.intersection(*groups)
    if not common:
        return None
    return max(common)


def infer_period_from_inputs(statement_period_groups, raw_period_groups):
    """Infer the target month without falling back to an older coincidental overlap."""
    statement_groups=[set(group) for group in statement_period_groups if group]
    raw_groups=[set(group) for group in raw_period_groups if group]

    statement_common=set.intersection(*statement_groups) if statement_groups else set()
    raw_periods=set().union(*raw_groups) if raw_groups else set()

    if statement_common:
        target=max(statement_common)
        if raw_periods:
            return (target,"statement+raw") if target in raw_periods else (None,"conflict")
        return target,"statement"
    if raw_periods:
        return max(raw_periods),"raw"
    return None,"unknown"


def period_alignment_evidence(prior_paths, douzone_paths, account_codes, period):
    """Return (exact code+amount matches, shared vendor codes) for a candidate month."""
    prior_pairs=Counter()
    raw_pairs=Counter()
    prior_codes=set()
    raw_codes=set()
    try:
        for path in prior_paths:
            rr=read_prior(path,Path(path).stem,period)
            for item in rr.items:
                prior_pairs[(item.vendor_code,item.amount)]+=1
                prior_codes.add(item.vendor_code)
        for path in douzone_paths:
            rr=read_douzone(path,account_codes or None,period)
            for item in rr.items:
                if item.debit>0:
                    raw_pairs[(item.vendor_code,item.debit)]+=1
                    raw_codes.add(item.vendor_code)
    except Exception:
        return 0,0

    exact=sum(min(count,raw_pairs.get(key,0)) for key,count in prior_pairs.items())
    shared_codes=len(prior_codes & raw_codes)
    return exact,shared_codes


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("명세서 대사")
        self.geometry("1120x780")
        self.minsize(900,650)
        self.configure(fg_color=BG)

        today=date.today()
        self.year=tk.IntVar(value=today.year)
        self.month=tk.IntVar(value=today.month)
        self.prior_paths=[]
        self.douzone_paths=[]
        self.account_codes=tk.StringVar(value="25301")
        self.period_text=tk.StringVar()
        self.status_text=tk.StringVar(value="대상 회계월과 파일을 선택한 뒤 대사를 시작하세요.")
        self.prior_count_text=tk.StringVar(value="선택된 파일 없음")
        self.raw_count_text=tk.StringVar(value="선택된 파일 없음")
        self.detail_hint_text=tk.StringVar(value="대사를 실행하면 검토가 필요한 항목을 여기에 보여줍니다.")

        self.results=[]
        self.new_items=[]
        self.issues=[]
        self.last_period=None
        self.last_source_paths=[]
        self.last_source_digests={}
        self.result_view="review"

        self._build()
        self._update_period()

    def _build(self):
        s=ttk.Style(self)
        s.theme_use("clam")
        s.configure(
            "TButton",
            font=("Segoe UI",10),
            padding=(14,9),
            background=CARD,
            foreground=TEXT,
            bordercolor=BORDER_STRONG,
            relief="flat",
        )
        s.map("TButton",background=[("active",SOFT)],foreground=[("disabled","#A1A1A6")])
        s.configure("TEntry",padding=8,fieldbackground=CARD,bordercolor=BORDER_STRONG)
        s.configure(
            "Accent.TButton",
            font=("Segoe UI Semibold",10),
            padding=(20,11),
            foreground="white",
            background=ACCENT,
            bordercolor=ACCENT,
            relief="flat",
        )
        s.map("Accent.TButton",background=[("active",ACCENT_HOVER),("disabled","#9DC7F5")])
        s.configure(
            "Segment.TButton",
            font=("Segoe UI Semibold",9),
            padding=(13,7),
            foreground=MUTED,
            background=SOFT,
            bordercolor=BORDER,
            relief="flat",
        )
        s.map("Segment.TButton",background=[("active","#EAEAEE")])
        s.configure(
            "SegmentActive.TButton",
            font=("Segoe UI Semibold",9),
            padding=(13,7),
            foreground=TEXT,
            background=CARD,
            bordercolor=BORDER_STRONG,
            relief="flat",
        )
        s.configure(
            "Treeview",
            font=("Segoe UI",9),
            rowheight=34,
            background=CARD,
            fieldbackground=CARD,
            foreground=TEXT,
            borderwidth=0,
        )
        s.map("Treeview",background=[("selected",SELECT_BG)],foreground=[("selected",TEXT)])
        s.configure(
            "Treeview.Heading",
            font=("Segoe UI Semibold",9),
            padding=(8,8),
            background=SOFT,
            foreground=MUTED,
            bordercolor=BORDER,
            relief="flat",
        )

        shell=tk.Frame(self,bg=BG)
        shell.pack(fill="both",expand=True)
        self.scroll_canvas=tk.Canvas(shell,bg=BG,highlightthickness=0,borderwidth=0)
        scroll=ttk.Scrollbar(shell,orient="vertical",command=self.scroll_canvas.yview)
        self.scroll_canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right",fill="y")
        self.scroll_canvas.pack(side="left",fill="both",expand=True)

        scroll_body=tk.Frame(self.scroll_canvas,bg=BG)
        self.scroll_window=self.scroll_canvas.create_window((0,0),window=scroll_body,anchor="nw")
        self.root_content=tk.Frame(scroll_body,bg=BG)
        self.root_content.pack(fill="both",expand=True,padx=44,pady=(30,38))
        root=self.root_content

        scroll_body.bind("<Configure>",lambda _e:self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all")))
        self.scroll_canvas.bind("<Configure>",self._on_canvas_configure)
        self.bind_all("<MouseWheel>",self._on_mousewheel,add="+")
        self.bind_all("<Button-4>",self._on_mousewheel,add="+")
        self.bind_all("<Button-5>",self._on_mousewheel,add="+")

        tk.Label(root,text="명세서 대사",font=("Segoe UI Semibold",30),bg=BG,fg=TEXT).pack(anchor="w")
        tk.Label(
            root,
            text="명세서와 더존 전표를 자동 대사하고, 사람이 확인할 항목만 남깁니다.",
            font=("Segoe UI",11),
            bg=BG,
            fg=MUTED,
        ).pack(anchor="w",pady=(5,22))

        period=ctk.CTkFrame(root,fg_color=CARD,border_width=1,border_color=BORDER,corner_radius=18)
        period.pack(fill="x",pady=(0,14))
        left=tk.Frame(period,bg=CARD)
        left.pack(side="left",padx=22,pady=18)
        tk.Label(left,text="대상 회계월",font=("Segoe UI Semibold",10),bg=CARD,fg=TEXT).pack(anchor="w")
        controls=tk.Frame(left,bg=CARD)
        controls.pack(anchor="w",pady=(8,0))
        self.year_spin=ttk.Spinbox(controls,from_=2020,to=2100,textvariable=self.year,width=7,state="readonly")
        self.year_spin.pack(side="left")
        tk.Label(controls,text="년",bg=CARD,fg=MUTED).pack(side="left",padx=(5,12))
        self.month_spin=ttk.Spinbox(controls,from_=1,to=12,textvariable=self.month,width=4,state="readonly")
        self.month_spin.pack(side="left")
        tk.Label(controls,text="월",bg=CARD,fg=MUTED).pack(side="left",padx=(5,10))
        tk.Label(controls,text="파일에서 자동 감지",font=("Segoe UI Semibold",9),bg=CARD,fg=GREEN).pack(side="left")
        self.year.trace_add("write",lambda *_:self._on_period_changed())
        self.month.trace_add("write",lambda *_:self._on_period_changed())
        tk.Label(
            period,
            textvariable=self.period_text,
            font=("Segoe UI Semibold",11),
            bg=CARD,
            fg=ACCENT,
            justify="left",
        ).pack(side="left",padx=(38,20))

        card=ctk.CTkFrame(root,fg_color=CARD,border_width=1,border_color=BORDER,corner_radius=20)
        card.pack(fill="x")
        card.grid_columnconfigure(0,minsize=165)
        card.grid_columnconfigure(1,weight=1)
        card.grid_columnconfigure(2,minsize=130)

        header=tk.Frame(card,bg=CARD)
        header.grid(row=0,column=0,columnspan=3,sticky="ew",padx=22,pady=(18,8))
        tk.Label(header,text="입력 파일",font=("Segoe UI Semibold",13),bg=CARD,fg=TEXT).pack(side="left")
        tk.Label(
            header,
            text="  원본은 수정하지 않고 읽기만 합니다.",
            font=("Segoe UI",9),
            bg=CARD,
            fg=MUTED,
        ).pack(side="left",pady=(2,0))

        prior_label=tk.Frame(card,bg=CARD)
        prior_label.grid(row=1,column=0,sticky="nw",padx=(22,12),pady=14)
        tk.Label(prior_label,text="명세서",font=("Segoe UI Semibold",10),bg=CARD,fg=TEXT).pack(anchor="w")
        tk.Label(prior_label,textvariable=self.prior_count_text,font=("Segoe UI",9),bg=CARD,fg=MUTED).pack(anchor="w",pady=(3,0))
        self.prior_list=tk.Listbox(
            card,height=3,font=("Segoe UI",9),selectmode="extended",
            bg=LIST_BG,fg=TEXT,selectbackground=SELECT_BG,selectforeground=TEXT,
            relief="flat",highlightthickness=1,highlightbackground=BORDER,highlightcolor=ACCENT,
        )
        self.prior_list.grid(row=1,column=1,sticky="ew",pady=14)
        pf=tk.Frame(card,bg=CARD)
        pf.grid(row=1,column=2,padx=18,pady=14,sticky="n")
        ctk.CTkButton(pf,text="파일 추가",command=self.pick_priors,width=108,height=34,corner_radius=11,
                      fg_color=SOFT,hover_color="#E9E9EE",text_color=TEXT).pack(fill="x")
        ctk.CTkButton(pf,text="선택 제거",command=self.remove_priors,width=108,height=34,corner_radius=11,
                      fg_color="transparent",hover_color=SOFT,text_color=MUTED,
                      border_width=1,border_color=BORDER).pack(fill="x",pady=(6,0))

        divider=tk.Frame(card,bg=BORDER,height=1)
        divider.grid(row=2,column=0,columnspan=3,sticky="ew",padx=22)

        raw_label=tk.Frame(card,bg=CARD)
        raw_label.grid(row=3,column=0,sticky="nw",padx=(22,12),pady=14)
        tk.Label(raw_label,text="더존 Raw",font=("Segoe UI Semibold",10),bg=CARD,fg=TEXT).pack(anchor="w")
        tk.Label(raw_label,textvariable=self.raw_count_text,font=("Segoe UI",9),bg=CARD,fg=MUTED).pack(anchor="w",pady=(3,0))
        self.douzone_list=tk.Listbox(
            card,height=3,font=("Segoe UI",9),selectmode="extended",
            bg=LIST_BG,fg=TEXT,selectbackground=SELECT_BG,selectforeground=TEXT,
            relief="flat",highlightthickness=1,highlightbackground=BORDER,highlightcolor=ACCENT,
        )
        self.douzone_list.grid(row=3,column=1,sticky="ew",pady=14)
        df=tk.Frame(card,bg=CARD)
        df.grid(row=3,column=2,padx=18,pady=14,sticky="n")
        ctk.CTkButton(df,text="파일 추가",command=self.pick_douzone,width=108,height=34,corner_radius=11,
                      fg_color=SOFT,hover_color="#E9E9EE",text_color=TEXT).pack(fill="x")
        ctk.CTkButton(df,text="선택 제거",command=self.remove_douzone,width=108,height=34,corner_radius=11,
                      fg_color="transparent",hover_color=SOFT,text_color=MUTED,
                      border_width=1,border_color=BORDER).pack(fill="x",pady=(6,0))

        options=tk.Frame(card,bg=CARD)
        options.grid(row=4,column=0,columnspan=3,sticky="ew",padx=22,pady=(4,10))
        tk.Label(options,text="미지급금 계정코드",font=("Segoe UI Semibold",10),bg=CARD,fg=TEXT).pack(side="left")
        ttk.Entry(options,textvariable=self.account_codes,width=16).pack(side="left",padx=(12,8))
        tk.Label(options,text="여러 개면 쉼표로 구분",font=("Segoe UI",9),bg=CARD,fg=MUTED).pack(side="left")

        tk.Label(
            card,
            text="명세서는 대상 회계월과 같은 월의 시트를 자동 선택합니다. 명세서 행 날짜는 장기이월 때문에 대사키로 사용하지 않습니다.",
            font=("Segoe UI",9),
            bg=CARD,
            fg=MUTED,
        ).grid(row=5,column=0,columnspan=3,sticky="w",padx=22,pady=(0,18))

        actions=tk.Frame(root,bg=BG)
        actions.pack(fill="x",pady=(16,12))
        ctk.CTkButton(
            actions,text="대사 시작",command=self.run,width=132,height=44,corner_radius=13,
            fg_color=ACCENT,hover_color=ACCENT_HOVER,text_color="white",
            font=("Segoe UI Semibold",10)
        ).pack(side="left")
        self.export_btn=ctk.CTkButton(
            actions,text="결과 Excel 저장",command=self.export,state="disabled",
            width=140,height=44,corner_radius=13,fg_color=CARD,hover_color=SOFT,
            text_color=TEXT,border_width=1,border_color=BORDER_STRONG,
            font=("Segoe UI Semibold",10)
        )
        self.export_btn.pack(side="left",padx=(10,0))

        self.banner=ctk.CTkFrame(
            root,fg_color=CARD,border_width=1,border_color=BORDER,corner_radius=18
        )
        self.banner.pack(fill="x",pady=(0,14))
        self.banner_title=ctk.CTkLabel(
            self.banner,text="대사 전",font=("Segoe UI Semibold",12),
            text_color=TEXT,anchor="w"
        )
        self.banner_title.pack(fill="x",padx=20,pady=(14,0))
        self.banner_detail=ctk.CTkLabel(
            self.banner,textvariable=self.status_text,font=("Segoe UI",9),
            text_color=MUTED,anchor="w"
        )
        self.banner_detail.pack(fill="x",padx=20,pady=(2,14))

        self.summary=tk.Frame(root,bg=BG)
        self.summary.pack(fill="x",pady=(0,14))
        self._summary({})

        self.preflight=tk.Label(
            root,text="",font=("Segoe UI",9),bg=BG,fg=WARN,justify="left",anchor="w"
        )
        self.preflight.pack(fill="x",pady=(0,10))

        result_card=ctk.CTkFrame(root,fg_color=CARD,border_width=1,border_color=BORDER,corner_radius=20)
        result_card.pack(fill="both",expand=True)

        result_header=tk.Frame(result_card,bg=CARD)
        result_header.pack(fill="x",padx=18,pady=(17,10))
        tk.Label(result_header,text="결과 상세",font=("Segoe UI Semibold",13),bg=CARD,fg=TEXT).pack(side="left")
        self.matched_view_btn=ctk.CTkButton(
            result_header,text="자동 대사 0",command=lambda:self._set_result_view("matched"),
            width=112,height=34,corner_radius=12,fg_color=SOFT,hover_color="#E9E9EE",
            text_color=MUTED
        )
        self.matched_view_btn.pack(side="right")
        self.review_view_btn=ctk.CTkButton(
            result_header,text="검토 필요 0",command=lambda:self._set_result_view("review"),
            width=112,height=34,corner_radius=12,fg_color=CARD,hover_color=SOFT,
            text_color=TEXT,border_width=1,border_color=BORDER_STRONG
        )
        self.review_view_btn.pack(side="right",padx=(0,6))

        tk.Label(
            result_card,textvariable=self.detail_hint_text,font=("Segoe UI",9),
            bg=CARD,fg=MUTED,anchor="w"
        ).pack(fill="x",padx=18,pady=(0,9))

        table_wrap=tk.Frame(result_card,bg=BORDER)
        table_wrap.pack(fill="both",expand=True,padx=18,pady=(0,18))
        self.detail=ttk.Treeview(
            table_wrap,
            columns=("owner","vendor","amount","status","reason"),
            show="headings",
            height=8,
        )
        columns=(
            ("owner","원본 명세",165),
            ("vendor","거래처",170),
            ("amount","금액",120),
            ("status","상태",140),
            ("reason","사유 / 참고",420),
        )
        for col,title,width in columns:
            self.detail.heading(col,text=title)
            self.detail.column(col,width=width,anchor="w")
        self.detail.tag_configure("matched",foreground=GREEN)
        self.detail.pack(fill="both",expand=True,padx=1,pady=1)

        tk.Label(
            root,
            text="안전 모드  ·  더존 직접 조작 없음  ·  날짜는 대사키로 사용하지 않음  ·  원본 덮어쓰기 금지",
            font=("Segoe UI",9),bg=BG,fg=MUTED,
        ).pack(anchor="w",pady=(14,0))

        self.account_codes.trace_add("write",lambda *_:self._on_account_code_changed())

    def _on_account_code_changed(self):
        self._invalidate_results()
        if self.prior_paths or self.douzone_paths:
            self.after_idle(lambda:self._maybe_align_period_from_inputs())

    def _on_canvas_configure(self,event):
        self.scroll_canvas.itemconfigure(self.scroll_window,width=event.width)
        side=max(32,(event.width-1240)//2)
        self.root_content.pack_configure(padx=side)

    def _on_mousewheel(self,event):
        if isinstance(event.widget,(tk.Listbox,ttk.Treeview)):
            return
        if getattr(event,"num",None)==4:
            steps=-1
        elif getattr(event,"num",None)==5:
            steps=1
        else:
            delta=getattr(event,"delta",0)
            if not delta:
                return
            steps=-1 if delta>0 else 1
        self.scroll_canvas.yview_scroll(steps,"units")

    def _set_banner(self,kind,title,detail):
        palette={
            "idle":(CARD,TEXT,MUTED,BORDER),
            "running":("#F2F7FF",ACCENT,MUTED,"#CFE2FF"),
            "success":(GREEN_BG,GREEN,MUTED,"#CFE8D9"),
            "warning":(WARN_BG,WARN,MUTED,"#F2D7A6"),
            "error":(RED_BG,RED,MUTED,"#F3C7C2"),
        }
        bg,title_fg,detail_fg,border=palette[kind]
        self.banner.configure(fg_color=bg,border_color=border)
        self.banner_title.configure(text=title,text_color=title_fg)
        self.banner_detail.configure(text_color=detail_fg)
        self.status_text.set(detail)

    def _update_period(self):
        try:
            p=AccountingPeriod(int(self.year.get()),int(self.month.get()))
            self.period_text.set(
                f"{p.label} 명세서  ↔  {p.label} 더존 전표\n"
                "명세서와 Raw에서 같은 회계월을 자동 감지합니다."
            )
        except Exception:
            self.period_text.set("올바른 연/월을 선택해주세요.")
        if hasattr(self,"prior_count_text"):
            self._refresh_file_counts()

    def _on_period_changed(self):
        self._update_period()
        self._invalidate_results()

    def _invalidate_results(self):
        if not hasattr(self,"export_btn"):
            return
        had_result=bool(self.results or self.new_items or self.issues or self.last_period)
        self.results=[]
        self.new_items=[]
        self.issues=[]
        self.last_period=None
        self.last_source_paths=[]
        self.last_source_digests={}
        self.export_btn.configure(state="disabled")
        self._summary({})
        self.preflight.config(text="")
        self._clear_detail()
        self.result_view="review"
        self._update_view_buttons()
        self.detail_hint_text.set("대사를 실행하면 검토가 필요한 항목을 여기에 보여줍니다.")
        if had_result:
            self._set_banner("warning","입력 조건 변경","입력 조건이 변경되었습니다. 다시 대사를 실행하세요.")
        else:
            self._set_banner("idle","대사 전","대상 회계월과 파일을 선택한 뒤 대사를 시작하세요.")

    def _refresh_file_counts(self):
        try:
            needed=AccountingPeriod(int(self.year.get()),int(self.month.get())).label
            prior_state=f"{len(self.prior_paths)}개 파일 선택" if self.prior_paths else "선택된 파일 없음"
            self.prior_count_text.set(f"필요: {needed} · {prior_state}")
        except Exception:
            self.prior_count_text.set(
                f"{len(self.prior_paths)}개 파일 선택" if self.prior_paths else "선택된 파일 없음"
            )
        self.raw_count_text.set(
            f"{len(self.douzone_paths)}개 파일 선택" if self.douzone_paths else "선택된 파일 없음"
        )

    def _maybe_align_period_from_inputs(self, show_banner=True):
        statement_groups=[]
        statement_labels=[]
        for path in self.prior_paths:
            try:
                periods=detect_statement_periods(path)
            except Exception:
                periods=[]
            if periods:
                statement_groups.append(periods)
                statement_labels.extend(periods)

        codes={x.strip() for x in self.account_codes.get().split(",") if x.strip()}
        raw_groups=[]
        raw_labels=[]
        for path in self.douzone_paths:
            try:
                periods=detect_douzone_periods(path,codes)
            except Exception:
                periods=[]
            if periods:
                raw_groups.append(periods)
                raw_labels.extend(periods)

        suggested,source=infer_period_from_inputs(statement_groups,raw_groups)
        if suggested is None:
            if source=="conflict":
                if show_banner:
                    statement_text=", ".join(p.label for p in sorted(set(statement_labels))) or "감지 안 됨"
                    raw_text=", ".join(p.label for p in sorted(set(raw_labels))) or "감지 안 됨"
                    self._set_banner(
                        "warning",
                        "회계월 자동 설정 보류",
                        f"명세서 월({statement_text})과 더존 Raw 월({raw_text})에서 같은 회계월을 찾지 못했습니다."
                    )
                return False
            return None

        if source=="statement+raw" and self.prior_paths and self.douzone_paths:
            exact,shared_codes=period_alignment_evidence(
                self.prior_paths,self.douzone_paths,codes,suggested
            )
            if exact==0 and shared_codes==0:
                if show_banner:
                    self._set_banner(
                        "warning",
                        "회계월 자동 설정 보류",
                        f"{suggested.label}이 양쪽 파일에 존재하지만 명세서와 Raw의 거래처코드가 하나도 겹치지 않습니다. "
                        "오래된 테스트 파일이나 다른 회계월 파일인지 확인해주세요."
                    )
                return False

        current=AccountingPeriod(int(self.year.get()),int(self.month.get()))
        if suggested!=current:
            # Set both variables without leaving a stale result behind.
            self.year.set(suggested.year)
            self.month.set(suggested.month)
            self._update_period()

        if show_banner:
            if source=="statement+raw":
                detail=f"{suggested.label} 명세서 + {suggested.label} 더존 Raw를 확인해 대상 회계월을 자동 설정했습니다."
            elif source=="statement":
                detail=f"명세서 월을 기준으로 대상 회계월을 {suggested.label}로 자동 설정했습니다."
            else:
                detail=f"더존 Raw의 미지급금 전표월을 기준으로 대상 회계월을 {suggested.label}로 자동 설정했습니다."
            self._refresh_file_counts()
            self._set_banner("idle","회계월 자동 설정",detail)
        return True

    def _add_classified_files(self,paths,requested_kind):
        """Add selected Excel files to the structurally correct input bucket."""
        added_prior=False
        added_raw=False
        moved=[]
        rejected=[]

        for p in paths:
            try:
                kind=classify_excel_input(p)
            except Exception as e:
                rejected.append(f"{Path(p).name}: 파일을 읽지 못함 ({e})")
                continue

            if kind=="prior":
                if p not in self.prior_paths:
                    self.prior_paths.append(p)
                    self.prior_list.insert("end",Path(p).name)
                    added_prior=True
                    if requested_kind!="prior":
                        moved.append(f"{Path(p).name} → 명세서")
            elif kind=="douzone":
                if p not in self.douzone_paths:
                    self.douzone_paths.append(p)
                    self.douzone_list.insert("end",Path(p).name)
                    added_raw=True
                    if requested_kind!="douzone":
                        moved.append(f"{Path(p).name} → 더존 Raw")
            elif kind=="ambiguous":
                rejected.append(f"{Path(p).name}: 명세서와 더존 구조가 함께 보여 자동 분류하지 않음")
            else:
                rejected.append(f"{Path(p).name}: 명세서/더존 Raw 구조를 식별하지 못함")

        if added_prior or added_raw:
            self._refresh_file_counts()
            self._invalidate_results()

        if added_prior or added_raw:
            self._maybe_align_period_from_inputs()

        if moved:
            self._set_banner(
                "idle",
                "파일 자동 분류",
                "잘못된 칸에서 선택한 파일을 구조에 맞게 자동 배치했습니다: " + " · ".join(moved)
            )

        if rejected:
            messagebox.showwarning(
                "파일 자동 분류 확인",
                "다음 파일은 안전하게 자동 분류할 수 없어 추가하지 않았습니다.\n\n"
                + "\n".join(rejected)
            )

    def pick_priors(self):
        paths=filedialog.askopenfilenames(filetypes=[("Excel","*.xlsx *.xlsm")])
        if paths:
            self._add_classified_files(paths,"prior")

    def pick_douzone(self):
        paths=filedialog.askopenfilenames(filetypes=[("Excel","*.xlsx *.xlsm")])
        if paths:
            self._add_classified_files(paths,"douzone")

    def remove_douzone(self):
        selected=list(self.douzone_list.curselection())
        for i in reversed(selected):
            self.douzone_list.delete(i)
            self.douzone_paths.pop(i)
        if selected:
            self._refresh_file_counts()
            self._invalidate_results()
            self._maybe_align_period_from_inputs()

    def remove_priors(self):
        selected=list(self.prior_list.curselection())
        for i in reversed(selected):
            self.prior_list.delete(i)
            self.prior_paths.pop(i)
        if selected:
            self._refresh_file_counts()
            self._invalidate_results()
            self._maybe_align_period_from_inputs()

    def _summary(self,counts):
        for w in self.summary.winfo_children():
            w.destroy()
        b=reconciliation_breakdown(counts,len(self.new_items),len(self.issues))
        cards=[
            ("자동 대사 완료",b["matched"],GREEN,"코드·금액 기준 자동 처리"),
            ("검토 필요",b["review"],WARN,"사람이 확인할 항목"),
            ("입력 확인",b["issues"],RED,"자동 처리에서 제외"),
            ("당월 신규 명세",b["new"],ACCENT,"현재는 검토 대상으로 유지"),
        ]
        for i,(name,value,color,sub) in enumerate(cards):
            box=ctk.CTkFrame(
                self.summary,fg_color=CARD,border_width=1,border_color=BORDER,corner_radius=17
            )
            box.grid(row=0,column=i,sticky="nsew",padx=(0 if i==0 else 7,0))
            self.summary.grid_columnconfigure(i,weight=1)
            ctk.CTkLabel(box,text=name,font=("Segoe UI Semibold",9),text_color=MUTED,anchor="w").pack(fill="x",padx=17,pady=(13,0))
            ctk.CTkLabel(box,text=f"{value:,}",font=("Segoe UI Semibold",25),text_color=color,anchor="w").pack(fill="x",padx=17,pady=(0,0))
            ctk.CTkLabel(box,text=sub,font=("Segoe UI",8),text_color=MUTED,anchor="w").pack(fill="x",padx=17,pady=(0,13))
        self._update_view_buttons()

    def _clear_detail(self):
        for row in self.detail.get_children():
            self.detail.delete(row)

    def _set_result_view(self,mode):
        self.result_view=mode
        self._update_view_buttons()
        self._refresh_detail()

    def _update_view_buttons(self):
        if not hasattr(self,"review_view_btn"):
            return
        matched=sum(1 for x in self.results if x.status==Status.MATCHED)
        review=sum(1 for x in self.results if x.status!=Status.MATCHED)
        review_active=self.result_view=="review"
        matched_active=self.result_view=="matched"
        self.review_view_btn.configure(
            text=f"검토 필요 {review:,}",
            fg_color=CARD if review_active else SOFT,
            text_color=TEXT if review_active else MUTED,
            border_width=1 if review_active else 0,
            border_color=BORDER_STRONG,
        )
        self.matched_view_btn.configure(
            text=f"자동 대사 {matched:,}",
            fg_color=CARD if matched_active else SOFT,
            text_color=TEXT if matched_active else MUTED,
            border_width=1 if matched_active else 0,
            border_color=BORDER_STRONG,
        )

    def _refresh_detail(self):
        self._clear_detail()
        if self.result_view=="matched":
            rows=[x for x in self.results if x.status==Status.MATCHED]
            self.detail_hint_text.set(
                "자동 대사된 항목입니다. 이름·적요 차이가 있었던 경우에도 참고 사유를 함께 남깁니다."
                if rows else "자동 대사된 항목이 없습니다."
            )
        else:
            rows=[x for x in self.results if x.status!=Status.MATCHED]
            self.detail_hint_text.set(
                "사람이 확인해야 하는 명세서 항목만 표시합니다."
                if rows else "검토할 명세서 항목이 없습니다."
            )
        for x in rows:
            tag="matched" if x.status==Status.MATCHED else ""
            self.detail.insert(
                "","end",
                values=(
                    x.prior.source.owner or x.prior.source.file_name,
                    x.prior.vendor_name,
                    f"{int(x.prior.amount):,}",
                    ("● " + x.status.value) if x.status==Status.MATCHED else x.status.value,
                    x.reason,
                ),
                tags=(tag,) if tag else (),
            )

    def run(self):
        if not self.prior_paths or not self.douzone_paths:
            messagebox.showwarning("파일 필요","명세서와 더존 Raw를 각각 1개 이상 추가해주세요.")
            return
        self._invalidate_results()
        sync_state=self._maybe_align_period_from_inputs(show_banner=False)
        if sync_state is False:
            self._set_banner("error","대사 중단","명세서 월과 더존 Raw 월이 서로 맞지 않습니다.")
            messagebox.showerror(
                "회계월 확인",
                "명세서와 더존 Raw에서 서로 연결되는 같은 회계월을 찾지 못했습니다.\n파일의 월을 확인해주세요."
            )
            return
        self._set_banner("running","대사 중","사전검사와 자동 대사를 진행하고 있습니다...")
        self.update_idletasks()
        try:
            source_paths=list(self.prior_paths)+list(self.douzone_paths)
            before_digests=snapshot_file_digests(source_paths)
            dup=find_duplicate_files(self.prior_paths,self.douzone_paths,before_digests)
            if dup:
                a,b=dup[0]
                raise ValueError(f"동일한 파일 내용이 중복 추가되었습니다: {a[1]} / {b[1]}")

            p=AccountingPeriod(int(self.year.get()),int(self.month.get()))
            codes={x.strip() for x in self.account_codes.get().split(",") if x.strip()}
            run=run_reconciliation(list(self.prior_paths),list(self.douzone_paths),codes,p)

            after_digests=snapshot_file_digests(source_paths)
            if before_digests != after_digests:
                raise ValueError("대사 실행 중 입력 Excel이 변경되었습니다. 파일 저장이 끝난 뒤 다시 실행해주세요.")

            self.issues=run.issues
            self.results=run.results
            self.new_items=run.new_items
            counts=run.counts
            self._summary(counts)

            breakdown=reconciliation_breakdown(counts,len(self.new_items),len(self.issues))
            self.preflight.config(
                text=(
                    f"입력 확인 {len(self.issues):,}건 · 해당 행은 자동 대사에서 제외했습니다."
                    if self.issues else
                    "사전검사 통과 · 자동 제외된 입력 이상이 없습니다."
                ),
                fg=(WARN if self.issues else GREEN),
            )

            self.result_view="review" if any(x.status!=Status.MATCHED for x in self.results) else "matched"
            self._update_view_buttons()
            self._refresh_detail()

            self.last_period=p
            self.last_source_paths=source_paths
            self.last_source_digests=after_digests

            detail=completion_detail(
                p,run.prior_count,counts,len(self.new_items),len(self.issues)
            )
            if breakdown["review"] or breakdown["issues"]:
                self._set_banner("success","대사 완료 · 확인 항목 있음",detail)
            else:
                self._set_banner("success","대사 완료",detail)
            self.export_btn.configure(state="normal")
        except Exception as e:
            self._set_banner("error","대사 중단","입력 내용을 확인한 뒤 다시 실행해주세요.")
            messagebox.showerror("대사 중단",str(e))

    def export(self):
        if self.last_period is None:
            messagebox.showwarning("대사 필요","현재 입력 조건으로 대사를 먼저 실행해주세요.")
            return
        changed=changed_snapshot_paths(self.last_source_digests)
        if changed:
            self._invalidate_results()
            messagebox.showwarning(
                "입력 파일 변경",
                "대사 후 입력 Excel이 변경되거나 사라졌습니다: "
                +", ".join(changed)+"\n다시 대사를 실행해주세요."
            )
            return

        p=self.last_period
        path=filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"{p.year}_{p.month:02d}_명세서_대사결과.xlsx",
            filetypes=[("Excel","*.xlsx")],
        )
        if not path:
            return
        try:
            write_result(path,self.results,self.new_items,self.issues,self.last_source_paths,p.label)
            self._set_banner(
                "success","결과 저장 완료",
                f"{Path(path).name} · 원본 Excel은 수정하지 않았습니다."
            )
            messagebox.showinfo(
                "저장 완료",
                "원본은 수정하지 않았습니다.\n확인필요/입력데이터확인 시트를 먼저 확인해주세요."
            )
        except Exception as e:
            messagebox.showerror("저장 실패",str(e))


if __name__=="__main__":
    App().mainloop()
