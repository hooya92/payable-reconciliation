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
MUTED="#86868B"
ACCENT="#0071E3"
ACCENT_HOVER="#0077ED"
GREEN="#248A3D"
GREEN_BG="#F0F9F2"
WARN="#A15C00"
WARN_BG="#FFF8E8"
RED="#B42318"
RED_BG="#FFF1F0"
BORDER="#E5E5EA"
BORDER_STRONG="#D2D2D7"
SOFT="#F5F5F7"
LIST_BG="#FBFBFD"
SELECT_BG="#E8F1FF"

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
        self.geometry("1180x820")
        self.minsize(980,700)
        self.configure(fg_color=BG)

        today=date.today()
        self.year=tk.IntVar(value=today.year)
        self.month=tk.IntVar(value=today.month)
        self.prior_paths=[]
        self.douzone_paths=[]
        self.account_codes=tk.StringVar(value="25301")
        self.period_text=tk.StringVar()
        self.period_badge_text=tk.StringVar(value="자동 감지 대기")
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
        # Native table styling kept for dense result inspection; the surrounding UI
        # uses rounded CustomTkinter surfaces and restrained Apple-like spacing.
        s=ttk.Style(self)
        s.theme_use("clam")
        s.configure(
            "Treeview",
            font=("Segoe UI",11),
            rowheight=40,
            background=CARD,
            fieldbackground=CARD,
            foreground=TEXT,
            borderwidth=0,
        )
        s.map("Treeview",background=[("selected",SELECT_BG)],foreground=[("selected",TEXT)])
        s.configure(
            "Treeview.Heading",
            font=("Segoe UI Semibold",10),
            padding=(10,10),
            background="#F7F7F9",
            foreground=MUTED,
            bordercolor=BORDER,
            relief="flat",
        )

        shell=ctk.CTkFrame(self,fg_color=BG,corner_radius=0)
        shell.pack(fill="both",expand=True)
        self.scroll_canvas=tk.Canvas(shell,bg=BG,highlightthickness=0,borderwidth=0)
        scroll=ctk.CTkScrollbar(
            shell,orientation="vertical",command=self.scroll_canvas.yview,
            width=12,fg_color=BG,button_color="#C7C7CC",button_hover_color="#AEAEB2"
        )
        self.scroll_canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right",fill="y",padx=(0,4),pady=6)
        self.scroll_canvas.pack(side="left",fill="both",expand=True)

        scroll_body=tk.Frame(self.scroll_canvas,bg=BG)
        self.scroll_window=self.scroll_canvas.create_window((0,0),window=scroll_body,anchor="nw")
        self.root_content=ctk.CTkFrame(scroll_body,fg_color="transparent",corner_radius=0)
        self.root_content.pack(fill="both",expand=True,padx=46,pady=(34,42))
        root=self.root_content

        scroll_body.bind("<Configure>",lambda _e:self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all")))
        self.scroll_canvas.bind("<Configure>",self._on_canvas_configure)
        self.bind_all("<MouseWheel>",self._on_mousewheel,add="+")
        self.bind_all("<Button-4>",self._on_mousewheel,add="+")
        self.bind_all("<Button-5>",self._on_mousewheel,add="+")

        # Hero
        hero=ctk.CTkFrame(root,fg_color="transparent")
        hero.pack(fill="x",pady=(0,22))
        hero.grid_columnconfigure(0,weight=1)
        ctk.CTkLabel(
            hero,text="명세서 대사",font=("Segoe UI Semibold",34),
            text_color=TEXT,anchor="w"
        ).grid(row=0,column=0,sticky="w")
        ctk.CTkLabel(
            hero,text="명세서와 더존 전표를 같은 회계월 기준으로 빠르게 대사합니다.",
            font=("Segoe UI",13),text_color=MUTED,anchor="w"
        ).grid(row=1,column=0,sticky="w",pady=(3,0))
        ctk.CTkLabel(
            hero,text="LOCAL  ·  READ ONLY",font=("Segoe UI Semibold",10),
            text_color=MUTED,fg_color="#ECECF0",corner_radius=11,
            padx=12,pady=5
        ).grid(row=0,column=1,rowspan=2,sticky="e")

        # Accounting period card
        period=ctk.CTkFrame(root,fg_color=CARD,border_width=1,border_color=BORDER,corner_radius=22)
        period.pack(fill="x",pady=(0,14))
        period.grid_columnconfigure(1,weight=1)

        period_left=ctk.CTkFrame(period,fg_color="transparent")
        period_left.grid(row=0,column=0,sticky="w",padx=(24,28),pady=20)
        ctk.CTkLabel(
            period_left,text="대상 회계월",font=("Segoe UI Semibold",11),
            text_color=MUTED,anchor="w"
        ).pack(anchor="w")
        badge_row=ctk.CTkFrame(period_left,fg_color="transparent")
        badge_row.pack(anchor="w",pady=(7,0))
        ctk.CTkLabel(
            badge_row,textvariable=self.period_badge_text,
            font=("Segoe UI Semibold",20),text_color=TEXT,
            fg_color="#F2F2F7",corner_radius=12,padx=14,pady=7
        ).pack(side="left")
        ctk.CTkLabel(
            badge_row,text="●  파일에서 자동 감지",
            font=("Segoe UI Semibold",10),text_color=GREEN,
            fg_color="#EDF8F0",corner_radius=11,padx=11,pady=5
        ).pack(side="left",padx=(9,0))

        ctk.CTkLabel(
            period,textvariable=self.period_text,
            font=("Segoe UI Semibold",12),text_color=ACCENT,
            justify="left",anchor="w"
        ).grid(row=0,column=1,sticky="w",padx=(12,24),pady=20)

        # Inputs
        card=ctk.CTkFrame(root,fg_color=CARD,border_width=1,border_color=BORDER,corner_radius=24)
        card.pack(fill="x")
        card.grid_columnconfigure(0,minsize=155)
        card.grid_columnconfigure(1,weight=1)
        card.grid_columnconfigure(2,minsize=124)

        header=ctk.CTkFrame(card,fg_color="transparent")
        header.grid(row=0,column=0,columnspan=3,sticky="ew",padx=24,pady=(22,10))
        ctk.CTkLabel(
            header,text="입력 파일",font=("Segoe UI Semibold",16),
            text_color=TEXT,anchor="w"
        ).pack(side="left")
        ctk.CTkLabel(
            header,text="원본 파일은 읽기만 하고 수정하지 않습니다.",
            font=("Segoe UI",10),text_color=MUTED,anchor="w"
        ).pack(side="left",padx=(12,0),pady=(2,0))

        def add_file_row(row,title,count_var,add_command,remove_command):
            label=ctk.CTkFrame(card,fg_color="transparent")
            label.grid(row=row,column=0,sticky="nw",padx=(24,12),pady=13)
            ctk.CTkLabel(
                label,text=title,font=("Segoe UI Semibold",11),
                text_color=TEXT,anchor="w"
            ).pack(anchor="w")
            ctk.CTkLabel(
                label,textvariable=count_var,font=("Segoe UI",9),
                text_color=MUTED,anchor="w"
            ).pack(anchor="w",pady=(3,0))

            list_wrap=ctk.CTkFrame(
                card,fg_color=LIST_BG,border_width=1,border_color=BORDER,
                corner_radius=14,height=58
            )
            list_wrap.grid(row=row,column=1,sticky="ew",pady=10)
            list_wrap.grid_propagate(False)
            lb=tk.Listbox(
                list_wrap,height=2,font=("Segoe UI",10),selectmode="extended",
                bg=LIST_BG,fg=TEXT,selectbackground=SELECT_BG,selectforeground=TEXT,
                relief="flat",highlightthickness=0,borderwidth=0,
                activestyle="none"
            )
            lb.pack(fill="both",expand=True,padx=10,pady=7)

            buttons=ctk.CTkFrame(card,fg_color="transparent")
            buttons.grid(row=row,column=2,padx=(16,18),pady=10,sticky="n")
            ctk.CTkButton(
                buttons,text="파일 추가",command=add_command,width=106,height=34,
                corner_radius=11,fg_color="#F0F0F4",hover_color="#E7E7EC",
                text_color=TEXT,font=("Segoe UI Semibold",10)
            ).pack(fill="x")
            ctk.CTkButton(
                buttons,text="선택 제거",command=remove_command,width=106,height=30,
                corner_radius=10,fg_color="transparent",hover_color=SOFT,
                text_color=MUTED,font=("Segoe UI",10)
            ).pack(fill="x",pady=(5,0))
            return lb

        self.prior_list=add_file_row(
            1,"명세서",self.prior_count_text,self.pick_priors,self.remove_priors
        )
        divider=ctk.CTkFrame(card,fg_color=BORDER,height=1,corner_radius=0)
        divider.grid(row=2,column=0,columnspan=3,sticky="ew",padx=24)
        self.douzone_list=add_file_row(
            3,"더존 Raw",self.raw_count_text,self.pick_douzone,self.remove_douzone
        )

        options=ctk.CTkFrame(card,fg_color="#F8F8FA",corner_radius=15)
        options.grid(row=4,column=0,columnspan=3,sticky="ew",padx=24,pady=(8,10))
        ctk.CTkLabel(
            options,text="미지급금 계정코드",font=("Segoe UI Semibold",11),
            text_color=TEXT
        ).pack(side="left",padx=(14,10),pady=11)
        ctk.CTkEntry(
            options,textvariable=self.account_codes,width=132,height=32,
            corner_radius=9,fg_color=CARD,border_color=BORDER_STRONG,
            text_color=TEXT,font=("Segoe UI",11)
        ).pack(side="left",pady=8)
        ctk.CTkLabel(
            options,text="여러 개면 쉼표로 구분",font=("Segoe UI",10),
            text_color=MUTED
        ).pack(side="left",padx=(10,0))

        ctk.CTkLabel(
            card,
            text="명세서는 대상 회계월과 같은 월의 시트를 자동 선택합니다.  ·  행 날짜는 장기이월 때문에 대사키로 사용하지 않습니다.",
            font=("Segoe UI",10),text_color=MUTED,anchor="w"
        ).grid(row=5,column=0,columnspan=3,sticky="ew",padx=24,pady=(0,20))

        # Primary actions
        actions=ctk.CTkFrame(root,fg_color="transparent")
        actions.pack(fill="x",pady=(16,12))
        ctk.CTkButton(
            actions,text="대사 시작",command=self.run,width=148,height=46,
            corner_radius=14,fg_color=ACCENT,hover_color=ACCENT_HOVER,
            text_color="white",font=("Segoe UI Semibold",12)
        ).pack(side="left")
        self.export_btn=ctk.CTkButton(
            actions,text="결과 Excel 저장",command=self.export,state="disabled",
            width=148,height=46,corner_radius=14,fg_color=CARD,hover_color="#FAFAFC",
            text_color=TEXT,border_width=1,border_color=BORDER_STRONG,
            font=("Segoe UI Semibold",11)
        )
        self.export_btn.pack(side="left",padx=(10,0))

        # Status
        self.banner=ctk.CTkFrame(
            root,fg_color=CARD,border_width=1,border_color=BORDER,corner_radius=20
        )
        self.banner.pack(fill="x",pady=(0,14))
        banner_inner=ctk.CTkFrame(self.banner,fg_color="transparent")
        banner_inner.pack(fill="x",padx=20,pady=14)
        self.banner_title=ctk.CTkLabel(
            banner_inner,text="대사 전",font=("Segoe UI Semibold",12),
            text_color=TEXT,anchor="w"
        )
        self.banner_title.pack(anchor="w")
        self.banner_detail=ctk.CTkLabel(
            banner_inner,textvariable=self.status_text,font=("Segoe UI",10),
            text_color=MUTED,anchor="w"
        )
        self.banner_detail.pack(fill="x",pady=(3,0))

        self.summary=ctk.CTkFrame(root,fg_color="transparent")
        self.summary.pack(fill="x",pady=(0,14))
        self._summary({})

        self.preflight=ctk.CTkLabel(
            root,text="",font=("Segoe UI",10),text_color=WARN,
            justify="left",anchor="w"
        )
        self.preflight.pack(fill="x",pady=(0,10))

        # Result card
        result_card=ctk.CTkFrame(root,fg_color=CARD,border_width=1,border_color=BORDER,corner_radius=24)
        result_card.pack(fill="both",expand=True)

        result_header=ctk.CTkFrame(result_card,fg_color="transparent")
        result_header.pack(fill="x",padx=20,pady=(18,8))
        ctk.CTkLabel(
            result_header,text="결과 상세",font=("Segoe UI Semibold",16),
            text_color=TEXT
        ).pack(side="left")
        self.matched_view_btn=ctk.CTkButton(
            result_header,text="자동 대사 0",command=lambda:self._set_result_view("matched"),
            width=112,height=34,corner_radius=11,fg_color=SOFT,hover_color="#ECECF0",
            text_color=MUTED,font=("Segoe UI Semibold",10)
        )
        self.matched_view_btn.pack(side="right")
        self.review_view_btn=ctk.CTkButton(
            result_header,text="검토 필요 0",command=lambda:self._set_result_view("review"),
            width=112,height=34,corner_radius=11,fg_color=CARD,hover_color=SOFT,
            text_color=TEXT,border_width=1,border_color=BORDER_STRONG,
            font=("Segoe UI Semibold",10)
        )
        self.review_view_btn.pack(side="right",padx=(0,6))

        ctk.CTkLabel(
            result_card,textvariable=self.detail_hint_text,font=("Segoe UI",10),
            text_color=MUTED,anchor="w"
        ).pack(fill="x",padx=20,pady=(0,10))

        table_wrap=ctk.CTkFrame(
            result_card,fg_color=BORDER,corner_radius=14,border_width=0
        )
        table_wrap.pack(fill="both",expand=True,padx=20,pady=(0,20))
        self.detail=ttk.Treeview(
            table_wrap,
            columns=("owner","vendor","amount","status","reason"),
            show="headings",height=8,
        )
        columns=(
            ("owner","원본 명세",165),
            ("vendor","거래처",180),
            ("amount","금액",120),
            ("status","상태",150),
            ("reason","사유 / 참고",430),
        )
        for col,title,width in columns:
            self.detail.heading(col,text=title)
            self.detail.column(col,width=width,anchor="w")
        self.detail.tag_configure("matched",foreground=GREEN)
        self.detail.pack(fill="both",expand=True,padx=1,pady=1)

        footer=ctk.CTkFrame(root,fg_color="#ECECF0",corner_radius=12)
        footer.pack(anchor="w",pady=(14,0))
        ctk.CTkLabel(
            footer,
            text="안전 모드  ·  더존 직접 조작 없음  ·  날짜는 대사키로 사용하지 않음  ·  원본 덮어쓰기 금지",
            font=("Segoe UI",10),text_color=MUTED
        ).pack(padx=12,pady=6)

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
            self.period_badge_text.set(p.label)
            self.period_text.set(
                f"{p.label} 명세서  ↔  {p.label} 더존 전표\n"
                "명세서와 Raw에서 같은 회계월을 자동 감지합니다."
            )
        except Exception:
            self.period_badge_text.set("자동 감지 대기")
            self.period_text.set("파일을 추가하면 회계월을 자동으로 감지합니다.")
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
        self.preflight.configure(text="")
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
            ("자동 대사",b["matched"],GREEN,"확정"),
            ("검토 필요",b["review"],WARN,"확인"),
            ("입력 확인",b["issues"],RED,"보류"),
            ("Raw 신규",b["new"],ACCENT,"검토"),
        ]
        for i,(name,value,color,badge) in enumerate(cards):
            box=ctk.CTkFrame(
                self.summary,fg_color=CARD,border_width=1,border_color=BORDER,
                corner_radius=20,height=102
            )
            box.grid(row=0,column=i,sticky="nsew",padx=(0 if i==0 else 7,0))
            box.grid_propagate(False)
            self.summary.grid_columnconfigure(i,weight=1)

            top=ctk.CTkFrame(box,fg_color="transparent")
            top.pack(fill="x",padx=16,pady=(14,0))
            ctk.CTkLabel(
                top,text=name,font=("Segoe UI Semibold",10),
                text_color=MUTED
            ).pack(side="left")
            ctk.CTkLabel(
                top,text=badge,font=("Segoe UI Semibold",9),
                text_color=color,fg_color="#F7F7F9",
                corner_radius=9,padx=8,pady=3
            ).pack(side="right")
            ctk.CTkLabel(
                box,text=f"{value:,}",font=("Segoe UI Semibold",29),
                text_color=TEXT,anchor="w"
            ).pack(fill="x",padx=16,pady=(2,12))
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
            self.preflight.configure(
                text=(
                    f"입력 확인 {len(self.issues):,}건 · 해당 행은 자동 대사에서 제외했습니다."
                    if self.issues else
                    "사전검사 통과 · 자동 제외된 입력 이상이 없습니다."
                ),
                text_color=(WARN if self.issues else GREEN),
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
