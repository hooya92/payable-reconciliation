from __future__ import annotations
import tkinter as tk
import customtkinter as ctk
import hashlib
import queue
import threading
import sys
from collections import Counter
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from adapters.excel.reader import classify_excel_input, detect_douzone_periods, detect_statement_periods, read_douzone, read_prior
from adapters.excel.writer import write_result
from adapters.excel.month_end_statement import write_month_end_statement
from domain.models import Status
from domain.period import AccountingPeriod
from application.service import run_reconciliation

BG="#F3F5F8"
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
SOFT="#F3F4F7"
LIST_BG="#FAFBFD"
SELECT_BG="#EAF2FF"

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
    auto_statuses=(Status.MATCHED,Status.SIGNED_NET_AUTO)
    matched=sum(counts.get(status,0) for status in auto_statuses)
    review=sum(v for k,v in counts.items() if k not in auto_statuses)+new_count
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
    return max(common).next()


def infer_period_from_inputs(statement_period_groups, raw_period_groups, preferred=None):
    """Infer the target month without falling back to an older coincidental overlap."""
    statement_groups=[set(group) for group in statement_period_groups if group]
    raw_groups=[set(group) for group in raw_period_groups if group]

    statement_common=set.intersection(*statement_groups) if statement_groups else set()
    raw_periods=set().union(*raw_groups) if raw_groups else set()

    if statement_common:
        target=max(statement_common).next()
        if preferred is not None and preferred.previous() in statement_common and (
            not raw_periods or preferred in raw_periods
        ):
            target=preferred
        if raw_periods:
            return (target,"statement+raw") if target in raw_periods else (None,"conflict")
        return target,"statement"
    if raw_periods:
        # A manually selected month must not be replaced by the latest month
        # in an annual Raw when the statement tab has no readable month.
        if preferred is not None and preferred in raw_periods:
            return preferred,"raw"
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
            rr=read_prior(path,Path(path).stem,period.previous())
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
        if sys.platform=="win32":
            self.iconbitmap(str(Path(__file__).resolve().parent/"assets"/"app.ico"))
        self.geometry("1280x920")
        self.minsize(1080,760)
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
        self.last_reconciliation_run=None
        self.last_period=None
        self.last_source_paths=[]
        self.last_source_digests={}
        self.result_view="review"
        self._file_add_busy=False
        self._file_add_queue=queue.Queue()

        self._build()
        self._update_period()

    def _build(self):
        s=ttk.Style(self)
        s.theme_use("clam")
        s.configure(
            "Treeview",
            font=("Segoe UI",12),
            rowheight=42,
            background=CARD,
            fieldbackground=CARD,
            foreground=TEXT,
            borderwidth=0,
        )
        s.map("Treeview",background=[("selected",SELECT_BG)],foreground=[("selected",TEXT)])
        s.configure(
            "Treeview.Heading",
            font=("Segoe UI Semibold",11),
            padding=(11,11),
            background="#F6F7F9",
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
        scroll.pack(side="right",fill="y",padx=(0,5),pady=8)
        self.scroll_canvas.pack(side="left",fill="both",expand=True)

        scroll_body=tk.Frame(self.scroll_canvas,bg=BG)
        self.scroll_window=self.scroll_canvas.create_window((0,0),window=scroll_body,anchor="nw")
        self.root_content=ctk.CTkFrame(scroll_body,fg_color="transparent",corner_radius=0)
        self.root_content.pack(fill="both",expand=True,padx=34,pady=(24,38))
        root=self.root_content

        scroll_body.bind("<Configure>",lambda _e:self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all")))
        self.scroll_canvas.bind("<Configure>",self._on_canvas_configure)
        self.bind_all("<MouseWheel>",self._on_mousewheel,add="+")
        self.bind_all("<Button-4>",self._on_mousewheel,add="+")
        self.bind_all("<Button-5>",self._on_mousewheel,add="+")

        # Hero card
        hero=ctk.CTkFrame(
            root,fg_color=CARD,border_width=1,border_color="#ECEEF2",
            corner_radius=24
        )
        hero.pack(fill="x",pady=(0,14))
        hero.grid_columnconfigure(0,weight=1)

        ctk.CTkLabel(
            hero,text="명세서 대사",font=("Segoe UI Semibold",34),
            text_color=TEXT,anchor="w"
        ).grid(row=0,column=0,sticky="sw",padx=(22,0),pady=(24,0))
        ctk.CTkLabel(
            hero,text="전월 말 명세서를 당월 더존 전표와 대사해 당월 말 명세서를 만듭니다.",
            font=("Segoe UI",13),text_color=MUTED,anchor="w"
        ).grid(row=1,column=0,sticky="nw",padx=(22,0),pady=(3,24))

        local_badge=ctk.CTkFrame(
            hero,fg_color="#F3F5F8",corner_radius=18
        )
        local_badge.grid(row=0,column=1,rowspan=2,sticky="e",padx=22)
        ctk.CTkLabel(
            local_badge,text="●",font=("Segoe UI",14),text_color="#34C759"
        ).pack(side="left",padx=(14,7),pady=9)
        ctk.CTkLabel(
            local_badge,text="LOCAL · READ ONLY",font=("Segoe UI Semibold",11),
            text_color="#667085"
        ).pack(side="left",padx=(0,14),pady=9)

        # Accounting period card
        period=ctk.CTkFrame(
            root,fg_color=CARD,border_width=1,border_color="#E6E8ED",
            corner_radius=22
        )
        period.pack(fill="x",pady=(0,14))
        period.grid_columnconfigure(1,weight=1)

        period_left=ctk.CTkFrame(period,fg_color="transparent")
        period_left.grid(row=0,column=0,sticky="w",padx=(22,20),pady=16)
        title_row=ctk.CTkFrame(period_left,fg_color="transparent")
        title_row.pack(anchor="w")
        ctk.CTkLabel(
            title_row,text="▣",font=("Segoe UI",16),text_color="#7890BB"
        ).pack(side="left",padx=(0,7))
        ctk.CTkLabel(
            title_row,text="대상 회계월",font=("Segoe UI Semibold",14),
            text_color=TEXT
        ).pack(side="left")

        badge_row=ctk.CTkFrame(period_left,fg_color="transparent")
        badge_row.pack(anchor="w",pady=(8,0))
        ctk.CTkLabel(
            badge_row,textvariable=self.period_badge_text,
            font=("Segoe UI Semibold",25),text_color=TEXT,
            fg_color="#F1F3F7",corner_radius=13,padx=16,pady=8
        ).pack(side="left")
        ctk.CTkLabel(
            badge_row,text="●  파일에서 자동 감지",
            font=("Segoe UI Semibold",12),text_color="#1F8F46",
            fg_color="#EAF8EE",corner_radius=13,padx=12,pady=7
        ).pack(side="left",padx=(10,0))

        divider=ctk.CTkFrame(period,width=1,height=54,fg_color="#E7E9EE",corner_radius=0)
        divider.grid(row=0,column=1,sticky="w",pady=20)
        ctk.CTkLabel(
            period,text="ⓘ",font=("Segoe UI Semibold",26),
            text_color=ACCENT
        ).grid(row=0,column=2,padx=(22,10))
        ctk.CTkLabel(
            period,textvariable=self.period_text,
            font=("Segoe UI",12),text_color=MUTED,
            justify="left",anchor="w"
        ).grid(row=0,column=3,sticky="w",padx=(0,24),pady=20)
        period.grid_columnconfigure(3,weight=1)

        # Input card
        card=ctk.CTkFrame(
            root,fg_color=CARD,border_width=1,border_color="#E6E8ED",
            corner_radius=22
        )
        card.pack(fill="x")
        card.grid_columnconfigure(0,minsize=150)
        card.grid_columnconfigure(1,weight=1)
        card.grid_columnconfigure(2,minsize=250)

        header=ctk.CTkFrame(card,fg_color="transparent")
        header.grid(row=0,column=0,columnspan=3,sticky="ew",padx=22,pady=(18,8))
        ctk.CTkLabel(
            header,text="▤",font=("Segoe UI",17),text_color="#7890BB"
        ).pack(side="left",padx=(0,8))
        ctk.CTkLabel(
            header,text="입력 파일",font=("Segoe UI Semibold",18),
            text_color=TEXT
        ).pack(side="left")
        ctk.CTkLabel(
            header,text="첨부 파일을 읽기만 하고 수정하지 않습니다.",
            font=("Segoe UI",12),text_color=MUTED
        ).pack(side="left",padx=(12,0),pady=(2,0))

        def add_file_row(row,title,count_var,add_command,remove_command):
            label=ctk.CTkFrame(card,fg_color="transparent")
            label.grid(row=row,column=0,sticky="nw",padx=(22,12),pady=12)
            ctk.CTkLabel(
                label,text=title,font=("Segoe UI Semibold",14),
                text_color=TEXT,anchor="w"
            ).pack(anchor="w")
            ctk.CTkLabel(
                label,textvariable=count_var,font=("Segoe UI",10),
                text_color=MUTED,anchor="w"
            ).pack(anchor="w",pady=(3,0))

            file_box=ctk.CTkFrame(
                card,fg_color="#FAFBFD",border_width=1,border_color="#DFE3EA",
                corner_radius=13,height=52
            )
            file_box.grid(row=row,column=1,sticky="ew",pady=10)
            file_box.grid_propagate(False)
            lb=tk.Listbox(
                file_box,height=1,font=("Segoe UI",12),selectmode="extended",
                bg="#FAFBFD",fg=TEXT,selectbackground=SELECT_BG,selectforeground=TEXT,
                relief="flat",highlightthickness=0,borderwidth=0,
                activestyle="none"
            )
            lb.pack(side="left",fill="both",expand=True,padx=(12,8),pady=8)

            actions=ctk.CTkFrame(card,fg_color="transparent")
            actions.grid(row=row,column=2,padx=(14,18),pady=10,sticky="e")
            ctk.CTkButton(
                actions,text="＋  파일 추가",command=add_command,width=124,height=42,
                corner_radius=11,fg_color="#EEF5FF",hover_color="#E4EFFF",
                text_color=ACCENT,font=("Segoe UI Semibold",12)
            ).pack(side="left")
            ctk.CTkButton(
                actions,text="첨부 제거",command=remove_command,width=116,height=42,
                corner_radius=11,fg_color="#F6F7F9",hover_color="#ECEEF2",
                text_color="#667085",font=("Segoe UI Semibold",12),
                border_width=1,border_color="#E2E5EA"
            ).pack(side="left",padx=(8,0))
            return lb

        self.prior_list=add_file_row(
            1,"명세서",self.prior_count_text,self.pick_priors,self.remove_priors
        )
        self.douzone_list=add_file_row(
            2,"더존 Raw",self.raw_count_text,self.pick_douzone,self.remove_douzone
        )

        options=ctk.CTkFrame(card,fg_color="#F7F8FA",corner_radius=14)
        options.grid(row=3,column=0,columnspan=3,sticky="ew",padx=22,pady=(8,12))
        ctk.CTkLabel(
            options,text="미지급금 계정코드",font=("Segoe UI Semibold",13),
            text_color=TEXT
        ).pack(side="left",padx=(14,10),pady=10)
        ctk.CTkEntry(
            options,textvariable=self.account_codes,width=170,height=38,
            corner_radius=9,fg_color=CARD,border_color="#D6DAE2",
            text_color=TEXT,font=("Segoe UI",12)
        ).pack(side="left",pady=8)
        ctk.CTkLabel(
            options,text="│  여러 개면 쉼표로 구분",font=("Segoe UI",12),
            text_color=MUTED
        ).pack(side="left",padx=(10,0))

        guide=ctk.CTkFrame(card,fg_color="#F4F6F9",corner_radius=12)
        guide.grid(row=4,column=0,columnspan=3,sticky="ew",padx=22,pady=(0,14))
        ctk.CTkLabel(
            guide,text="ⓘ",font=("Segoe UI Semibold",16),text_color=ACCENT
        ).pack(side="left",padx=(12,8),pady=9)
        ctk.CTkLabel(
            guide,
            text="명세서는 대상 회계월의 전월 말 시트를 선택합니다.  ·  행 날짜는 장기이월 때문에 대사키로 사용하지 않습니다.",
            font=("Segoe UI",11),text_color="#667085",anchor="w"
        ).pack(side="left",fill="x",expand=True,padx=(0,12),pady=9)

        # Primary actions
        actions=ctk.CTkFrame(root,fg_color="transparent")
        actions.pack(fill="x",pady=(14,10))
        ctk.CTkButton(
            actions,text="▶   대사 시작",command=self.run,width=178,height=52,
            corner_radius=13,fg_color=ACCENT,hover_color=ACCENT_HOVER,
            text_color="white",font=("Segoe UI Semibold",14)
        ).pack(side="left")
        self.month_end_btn=ctk.CTkButton(
            actions,text="▤   당월 명세서 생성",command=self.export_month_end,state="disabled",
            width=198,height=52,corner_radius=13,fg_color="#EDE7F6",hover_color="#E4D8F5",
            text_color="#6040A0",font=("Segoe UI Semibold",13)
        )
        self.month_end_btn.pack(side="left",padx=(10,0))

        # Status
        self.banner=ctk.CTkFrame(
            root,fg_color=CARD,border_width=1,border_color=BORDER,corner_radius=20
        )
        self.banner.pack(fill="x",pady=(0,12))
        banner_inner=ctk.CTkFrame(self.banner,fg_color="transparent")
        banner_inner.pack(fill="x",padx=18,pady=13)
        self.banner_icon=ctk.CTkLabel(
            banner_inner,text="●",font=("Segoe UI Semibold",20),
            text_color=GREEN
        )
        self.banner_icon.pack(side="left",padx=(0,10))
        banner_text=ctk.CTkFrame(banner_inner,fg_color="transparent")
        banner_text.pack(side="left",fill="x",expand=True)
        self.banner_title=ctk.CTkLabel(
            banner_text,text="대사 전",font=("Segoe UI Semibold",17),
            text_color=TEXT,anchor="w"
        )
        self.banner_title.pack(anchor="w")
        self.banner_detail=ctk.CTkLabel(
            banner_text,textvariable=self.status_text,font=("Segoe UI",12),
            text_color=MUTED,anchor="w"
        )
        self.banner_detail.pack(fill="x",pady=(2,0))

        self.summary=ctk.CTkFrame(root,fg_color="transparent")
        self.summary.pack(fill="x",pady=(0,12))
        self._summary({})

        self.preflight=ctk.CTkLabel(
            root,text="",font=("Segoe UI",10),text_color=WARN,
            justify="left",anchor="w"
        )
        self.preflight.pack(fill="x",pady=(0,8))

        # Result card
        result_card=ctk.CTkFrame(
            root,fg_color=CARD,border_width=1,border_color="#E6E8ED",
            corner_radius=22
        )
        result_card.pack(fill="both",expand=True)

        result_header=ctk.CTkFrame(result_card,fg_color="transparent")
        result_header.pack(fill="x",padx=18,pady=(16,8))
        ctk.CTkLabel(
            result_header,text="☷",font=("Segoe UI",17),text_color="#7890BB"
        ).pack(side="left",padx=(0,8))
        ctk.CTkLabel(
            result_header,text="결과 상세",font=("Segoe UI Semibold",17),
            text_color=TEXT
        ).pack(side="left")

        self.matched_view_btn=ctk.CTkButton(
            result_header,text="자동 대사 0",command=lambda:self._set_result_view("matched"),
            width=132,height=40,corner_radius=14,fg_color="#F3F4F7",hover_color="#ECEEF2",
            text_color=MUTED,font=("Segoe UI Semibold",12)
        )
        self.matched_view_btn.pack(side="right")
        self.review_view_btn=ctk.CTkButton(
            result_header,text="검토 필요 0",command=lambda:self._set_result_view("review"),
            width=132,height=40,corner_radius=14,fg_color="#F3F4F7",hover_color="#ECEEF2",
            text_color=MUTED,font=("Segoe UI Semibold",12)
        )
        self.review_view_btn.pack(side="right",padx=(0,6))
        self.all_view_btn=ctk.CTkButton(
            result_header,text="전체",command=lambda:self._set_result_view("all"),
            width=96,height=40,corner_radius=14,fg_color=ACCENT,hover_color=ACCENT_HOVER,
            text_color="white",font=("Segoe UI Semibold",12)
        )
        self.all_view_btn.pack(side="right",padx=(0,6))

        ctk.CTkLabel(
            result_card,textvariable=self.detail_hint_text,font=("Segoe UI",11),
            text_color=MUTED,anchor="w"
        ).pack(fill="x",padx=18,pady=(0,8))

        table_wrap=ctk.CTkFrame(
            result_card,fg_color="#E4E7EC",corner_radius=12,border_width=0
        )
        table_wrap.pack(fill="both",expand=True,padx=18,pady=(0,18))
        self.detail=ttk.Treeview(
            table_wrap,
            columns=("owner","vendor","amount","status","reason"),
            show="headings",height=8,
        )
        columns=(
            ("owner","원본 명세",175),
            ("vendor","거래처",190),
            ("amount","금액",120),
            ("status","상태",145),
            ("reason","사유 / 참고",430),
        )
        for col,title,width in columns:
            self.detail.heading(col,text=title)
            self.detail.column(col,width=width,anchor="w")
        self.detail.tag_configure("matched",foreground="#248A3D")
        self.detail.tag_configure("review",foreground="#8A5A00")
        self.detail.pack(fill="both",expand=True,padx=1,pady=1)

        footer=ctk.CTkFrame(root,fg_color="#ECEEF2",corner_radius=11)
        footer.pack(anchor="w",pady=(12,0))
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
        icon={"idle":"●","running":"●","success":"✓","warning":"!","error":"×"}[kind]
        icon_color={"idle":MUTED,"running":ACCENT,"success":GREEN,"warning":WARN,"error":RED}[kind]
        if hasattr(self,"banner_icon"):
            self.banner_icon.configure(text=icon,text_color=icon_color)
        self.banner.configure(fg_color=bg,border_color=border)
        self.banner_title.configure(text=title,text_color=title_fg)
        self.banner_detail.configure(text_color=detail_fg)
        self.status_text.set(detail)

    def _update_period(self):
        try:
            p=AccountingPeriod(int(self.year.get()),int(self.month.get()))
            self.period_badge_text.set(p.label)
            self.period_text.set(
                f"{p.label} 선택 → {p.previous().label} 말 명세서 + {p.label} 더존 원장\n"
                f"{p.label} 말 미지급금 명세서 초안 생성"
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
        if not hasattr(self,"month_end_btn"):
            return
        had_result=bool(self.results or self.new_items or self.issues or self.last_period)
        self.results=[]
        self.new_items=[]
        self.issues=[]
        self.last_reconciliation_run=None
        self.last_period=None
        self.last_source_paths=[]
        self.last_source_digests={}
        self.month_end_btn.configure(state="disabled")
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
        if not self.prior_paths and not self.douzone_paths:
            self.prior_count_text.set("선택된 파일 없음")
            self.raw_count_text.set("선택된 파일 없음")
            return
        try:
            needed=AccountingPeriod(int(self.year.get()),int(self.month.get())).previous().label
            prior_state=f"{len(self.prior_paths)}개 파일 선택" if self.prior_paths else "선택된 파일 없음"
            self.prior_count_text.set(f"필요: {needed} · {prior_state}" if self.prior_paths else prior_state)
        except Exception:
            self.prior_count_text.set(
                f"{len(self.prior_paths)}개 파일 선택" if self.prior_paths else "선택된 파일 없음"
            )
        self.raw_count_text.set(
            f"{len(self.douzone_paths)}개 파일 선택" if self.douzone_paths else "선택된 파일 없음"
        )

    def _maybe_align_period_from_inputs(self, show_banner=True, detected=None):
        if not self.prior_paths and not self.douzone_paths:
            self.period_badge_text.set("자동 감지 대기")
            self.period_text.set("명세서와 Raw 파일을 추가하면 회계월을 자동으로 감지합니다.")
            self._refresh_file_counts()
            if show_banner:
                self._set_banner("idle","대사 전","명세서와 더존 Raw 파일을 추가해주세요.")
            return None

        statement_groups=[]
        statement_labels=[]
        for path in self.prior_paths:
            try:
                periods=(detected[path] if detected is not None and path in detected
                         else detect_statement_periods(path))
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
                periods=(detected[path] if detected is not None and path in detected
                         else detect_douzone_periods(path,codes))
            except Exception:
                periods=[]
            if periods:
                raw_groups.append(periods)
                raw_labels.extend(periods)

        current=AccountingPeriod(int(self.year.get()),int(self.month.get()))
        suggested,source=infer_period_from_inputs(statement_groups,raw_groups,preferred=current)
        if suggested is None:
            if source=="conflict":
                if show_banner:
                    statement_text=", ".join(p.label for p in sorted(set(statement_labels))) or "감지 안 됨"
                    raw_text=", ".join(p.label for p in sorted(set(raw_labels))) or "감지 안 됨"
                    self._set_banner(
                        "warning",
                        "회계월 자동 설정 보류",
                        f"전월 명세서({statement_text})와 당월 더존 Raw({raw_text})가 이어지지 않습니다."
                    )
                return False
            return None

        if suggested!=current:
            # Set both variables without leaving a stale result behind.
            self.year.set(suggested.year)
            self.month.set(suggested.month)
            self._update_period()

        if show_banner:
            if source=="statement+raw":
                detail=f"{suggested.previous().label} 명세서 + {suggested.label} 더존 Raw를 확인해 대상 회계월을 자동 설정했습니다."
            elif source=="statement":
                detail=f"전월 명세서 기준으로 당월 회계월을 {suggested.label}로 자동 설정했습니다."
            else:
                detail=f"더존 Raw의 미지급금 전표월을 기준으로 대상 회계월을 {suggested.label}로 자동 설정했습니다."
            self._refresh_file_counts()
            self._set_banner("idle","회계월 자동 설정",detail)
        return True

    def _add_classified_files(self,paths,requested_kind):
        """Read and classify large files off the Tk main thread."""
        if self._file_add_busy:
            self._set_banner("warning","파일 분석 중","현재 파일 분석이 끝난 뒤 다시 추가해주세요.")
            return
        self._file_add_busy=True
        # Show selection immediately; this is a temporary label, not an accepted input.
        self._pending_file_list = self.prior_list if requested_kind == "prior" else self.douzone_list
        self._pending_file_labels = [f"분석 중 · {Path(p).name}" for p in paths]
        for label in self._pending_file_labels:
            self._pending_file_list.insert("end", label)
        self._set_banner("running","파일 분석 중","Excel 구조와 회계월을 확인하고 있습니다. 큰 파일은 시간이 걸릴 수 있습니다.")
        codes={x.strip() for x in self.account_codes.get().split(",") if x.strip()}
        prior=tuple(self.prior_paths)
        raw=tuple(self.douzone_paths)
        def worker():
            classified=[]
            detected={}
            try:
                for p in paths:
                    try:
                        kind=classify_excel_input(p)
                        classified.append((p,kind,None))
                    except Exception as e:
                        classified.append((p,None,str(e)))
                # Include existing files; period detection must use the same inputs as before.
                new_prior=prior+tuple(p for p,k,e in classified if k=="prior" and not e and p not in prior)
                new_raw=raw+tuple(p for p,k,e in classified if k=="douzone" and not e and p not in raw)
                for p in new_prior:
                    try: detected[p]=detect_statement_periods(p)
                    except Exception: detected[p]=[]
                for p in new_raw:
                    try: detected[p]=detect_douzone_periods(p,codes)
                    except Exception: detected[p]=[]
                self._file_add_queue.put((classified,requested_kind,detected,None))
            except Exception as e:
                self._file_add_queue.put((None,None,None,str(e)))
        threading.Thread(target=worker,daemon=True).start()
        self.after(100,self._poll_file_add)

    def _poll_file_add(self):
        try:
            classified,kind,detected,error=self._file_add_queue.get_nowait()
        except queue.Empty:
            self.after(250,self._poll_file_add)
            return
        try:
            # Remove only temporary rows before showing verified file names.
            for label in self._pending_file_labels:
                rows=self._pending_file_list.get(0,"end")
                if label in rows:
                    self._pending_file_list.delete(rows.index(label))
            if error:
                self._set_banner("error","파일 분석 실패",error)
                messagebox.showerror("파일 분석 실패",error)
            else:
                self._finish_classified_files(classified,kind,detected)
        finally:
            self._pending_file_labels=[]
            self._file_add_busy=False

    def _finish_classified_files(self,classified,requested_kind,detected):
        """Add selected Excel files to the structurally correct input bucket."""
        added_prior=False
        added_raw=False
        moved=[]
        rejected=[]

        for p,kind,error in classified:
            if error:
                rejected.append(f"{Path(p).name}: 파일을 읽지 못함 ({error})")
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
            self._maybe_align_period_from_inputs(detected=detected)

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
        paths=filedialog.askopenfilenames(filetypes=[("Excel","*.xlsx *.xlsm *.xls")])
        if paths:
            self._add_classified_files(paths,"prior")

    def pick_douzone(self):
        paths=filedialog.askopenfilenames(filetypes=[("Excel","*.xlsx *.xlsm *.xls")])
        if paths:
            self._add_classified_files(paths,"douzone")

    def remove_douzone(self):
        if self._file_add_busy:
            return
        if not self.douzone_paths:
            return
        self.douzone_list.delete(0,"end")
        self.douzone_paths.clear()
        self._refresh_file_counts()
        self._invalidate_results()
        self._maybe_align_period_from_inputs()

    def remove_priors(self):
        if self._file_add_busy:
            return
        if not self.prior_paths:
            return
        self.prior_list.delete(0,"end")
        self.prior_paths.clear()
        self._refresh_file_counts()
        self._invalidate_results()
        self._maybe_align_period_from_inputs()

    def _summary(self,counts):
        for w in self.summary.winfo_children():
            w.destroy()
        b=reconciliation_breakdown(counts,len(self.new_items),len(self.issues))
        cards=[
            ("자동 대사 완료",b["matched"],"#EAF8EE",GREEN,"✓"),
            ("검토 필요",b["review"],"#FFF4D8","#E58A00","!"),
            ("입력 확인",b["issues"],"#EAF3FF",ACCENT,"▤"),
            ("Raw 신규",b["new"],"#F4EAFE","#7A4ED6","+"),
        ]
        for i,(name,value,bg,color,icon) in enumerate(cards):
            box=ctk.CTkFrame(
                self.summary,fg_color=bg,border_width=1,border_color=BORDER,
                corner_radius=17,height=88
            )
            box.grid(row=0,column=i,sticky="nsew",padx=(0 if i==0 else 8,0))
            box.grid_propagate(False)
            self.summary.grid_columnconfigure(i,weight=1)

            icon_box=ctk.CTkLabel(
                box,text=icon,font=("Segoe UI Semibold",16),
                text_color=color,fg_color=CARD,width=40,height=40,
                corner_radius=20
            )
            icon_box.pack(side="left",padx=(14,10),pady=15)

            body=ctk.CTkFrame(box,fg_color="transparent")
            body.pack(side="left",fill="both",expand=True,pady=(12,10))
            ctk.CTkLabel(
                body,text=name,font=("Segoe UI Semibold",12),
                text_color="#344054",anchor="w"
            ).pack(anchor="w")
            ctk.CTkLabel(
                body,text=f"{value:,}건",font=("Segoe UI Semibold",24),
                text_color=TEXT,anchor="w"
            ).pack(anchor="w",pady=(1,0))
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
        matched=sum(1 for x in self.results if x.status in (Status.MATCHED,Status.SIGNED_NET_AUTO))
        review=sum(1 for x in self.results if x.status not in (Status.MATCHED,Status.SIGNED_NET_AUTO))

        def style(btn,active,label):
            btn.configure(
                text=label,
                fg_color=ACCENT if active else "#F3F4F7",
                text_color="white" if active else MUTED,
                border_width=0,
            )

        style(self.all_view_btn,self.result_view=="all","전체")
        style(self.review_view_btn,self.result_view=="review",f"검토 필요 {review:,}")
        style(self.matched_view_btn,self.result_view=="matched",f"자동 대사 {matched:,}")

    def _refresh_detail(self):
        self._clear_detail()
        if self.result_view=="matched":
            rows=[x for x in self.results if x.status in (Status.MATCHED,Status.SIGNED_NET_AUTO)]
            self.detail_hint_text.set(
                "자동 대사된 항목입니다. 이름·적요 차이가 있었던 경우에도 참고 사유를 함께 남깁니다."
                if rows else "자동 대사된 항목이 없습니다."
            )
        elif self.result_view=="review":
            rows=[x for x in self.results if x.status not in (Status.MATCHED,Status.SIGNED_NET_AUTO)]
            self.detail_hint_text.set(
                "사람이 확인해야 하는 명세서 항목만 표시합니다."
                if rows else "검토할 명세서 항목이 없습니다."
            )
        else:
            rows=list(self.results)
            self.detail_hint_text.set(
                "대사 결과 전체를 표시합니다."
                if rows else "대사 결과가 없습니다."
            )

        for x in rows:
            tag="matched" if x.status in (Status.MATCHED,Status.SIGNED_NET_AUTO) else "review"
            self.detail.insert(
                "","end",
                values=(
                    x.prior.source.owner or x.prior.source.file_name,
                    x.prior.vendor_name,
                    f"{int(x.closing_balance if x.status==Status.SIGNED_NET_AUTO else x.prior.amount):,}",
                    ("● " + x.status.value) if x.status in (Status.MATCHED,Status.SIGNED_NET_AUTO) else x.status.value,
                    x.reason,
                ),
                tags=(tag,),
            )

    def run(self):
        if self._file_add_busy:
            messagebox.showinfo("파일 분석 중","파일 분석이 완료된 후 대사를 실행해주세요.")
            return
        if not self.prior_paths or not self.douzone_paths:
            messagebox.showwarning("파일 필요","명세서와 더존 Raw를 각각 1개 이상 추가해주세요.")
            return
        self._invalidate_results()
        sync_state=self._maybe_align_period_from_inputs(show_banner=False)
        if sync_state is False:
            self._set_banner("error","대사 중단","전월 명세서와 당월 더존 Raw의 회계월이 맞지 않습니다.")
            messagebox.showerror(
                "회계월 확인",
                "전월 말 명세서와 당월 더존 Raw가 이어지는 월인지 확인해주세요.\n파일의 월을 확인해주세요."
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
            self.last_reconciliation_run=run
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

            self.result_view="review" if any(x.status not in (Status.MATCHED,Status.SIGNED_NET_AUTO) for x in self.results) else "matched"
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
            self.month_end_btn.configure(state="normal")
        except Exception as e:
            self._set_banner("error","대사 중단","입력 내용을 확인한 뒤 다시 실행해주세요.")
            messagebox.showerror("대사 중단",str(e))

    def export_month_end(self):
        """Save a separate, editable draft based on the original statement template."""
        if self.last_period is None or self.last_reconciliation_run is None:
            messagebox.showwarning("대사 필요","대사를 먼저 실행해주세요.")
            return
        changed=changed_snapshot_paths(self.last_source_digests)
        if changed:
            self._invalidate_results()
            messagebox.showwarning(
                "입력 파일 변경",
                "대사 후 원본 Excel이 변경되거나 사라졌습니다: "
                + ", ".join(changed) + "\n다시 대사를 실행해주세요."
            )
            return
        p=self.last_period
        path=filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"{p.year}_{p.month:02d}_당월명세서_검토용.xlsx",
            filetypes=[("Excel","*.xlsx")],
        )
        if not path:
            return
        try:
            run=self.last_reconciliation_run
            write_month_end_statement(
                path,self.prior_paths,run.results,run.new_items,run.issues,p,
                standalone_debits=run.standalone_debits,
                standalone_debit_sources=run.standalone_debit_sources,
                journal_items=run.journal_items,
                journal_sources=run.journal_sources,
                source_paths=self.last_source_paths,
            )
            self._set_banner(
                "success","당월 명세서 초안 생성",
                f"{Path(path).name} · 전월 양식 기반, 검토필요·변경내역 시트는 삭제 가능합니다."
            )
        except PermissionError as e:
            messagebox.showerror(
                "파일 접근 거부",
                "저장할 결과 파일 또는 관련 Excel 파일이 열려 있어 접근이 거부됐을 수 있습니다.\n"
                "Excel에서 해당 파일을 닫은 뒤 다시 저장해 주세요.\n\n"
                "계속 실패하면 다른 파일명이나 쓰기 가능한 폴더에 저장해 주세요.\n"
                "파일·폴더의 읽기 전용 설정과 쓰기 권한도 확인해 주세요.\n\n"
                f"문제 파일: {e.filename or path}"
            )
        except Exception as e:
            messagebox.showerror("당월 명세서 생성 실패",str(e))

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
        except Exception as e:
            messagebox.showerror("저장 실패",str(e))


if __name__=="__main__":
    App().mainloop()
