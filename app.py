from __future__ import annotations
import tkinter as tk
import customtkinter as ctk
import hashlib
import os
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
from domain.normalization import parse_account_codes
from domain.period import AccountingPeriod
from application.service import run_reconciliation
from adapters import updater
from application.workflow import (
    file_digest,
    open_saved_folder,
    snapshot_file_digests,
    changed_snapshot_paths,
    find_duplicate_files,
    reconciliation_breakdown,
    completion_detail,
    cached_periods,
    suggest_reconciliation_period,
    infer_period_from_inputs,
    period_alignment_evidence,
)

# Colour tokens. Every text/background pair below is >= 4.5:1 (WCAG AA).
# Light grey ground, white groups, one blue for actions; green/amber/red mean state only.
BG="#F5F5F7"
CARD="#FFFFFF"
TEXT="#1D1D1F"
TEXT2="#424245"
MUTED="#636366"
DISABLED_FG="#6B6B70"
ACCENT="#0066CC"
ACCENT_HOVER="#0058B0"
GREEN="#1B7233"
WARN="#8A5A00"
RED="#B42318"
BORDER="#E5E5EA"
SOFT="#EDEDF0"
SOFT_HOVER="#E1E1E6"
SELECT_BG="#E4EEFB"

# Segoe UI has no Hangul glyphs; Windows substitutes a regular-weight font, so
# bold headings would silently lose their weight. Use the Korean UI font directly.
FONT="Malgun Gothic" if sys.platform=="win32" else "Segoe UI"


BUTTON_PRIMARY=dict(fg_color=ACCENT,hover_color=ACCENT_HOVER,text_color="white",text_color_disabled="white")
BUTTON_SECONDARY=dict(fg_color=SOFT,hover_color=SOFT_HOVER,text_color=TEXT,text_color_disabled=DISABLED_FG)


def font(size, bold=False):
    return (FONT,size,"bold") if bold else (FONT,size)


IDLE_HINT="파일 추가 → 대사 시작 → 결과 검토 후 '당월 명세서 생성'"


ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("명세서 대사")
        if sys.platform=="win32":
            self.iconbitmap(str(Path(__file__).resolve().parent/"assets"/"app.ico"))
        self.geometry("500x800")
        self.minsize(480,700)
        self.configure(fg_color=BG)

        today=date.today()
        self.year=tk.IntVar(value=today.year)
        self.month=tk.IntVar(value=today.month)
        self.prior_paths=[]
        self.douzone_paths=[]
        self.account_codes=tk.StringVar(value="25301")
        self.period_text=tk.StringVar()
        self.period_value_text=tk.StringVar(value="")
        self._period_picker_open=False
        self._detect_state="none"
        self.year_text=tk.StringVar(value=str(today.year))
        self.month_text=tk.StringVar(value=str(today.month))
        self.status_text=tk.StringVar(value=IDLE_HINT)
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
        self._period_manual=False
        self._run_busy=False
        self._run_queue=queue.Queue()
        self._detail_reasons={}
        self._file_add_busy=False
        self._file_add_queue=queue.Queue()
        self._period_cache={}
        self._update_queue=queue.Queue()
        self._update_release=None
        self._update_text="새 버전 업데이트"

        self._build()
        self._update_period()
        self._sync_file_hints()
        if updater.can_self_update():
            threading.Thread(target=updater.cleanup_stale,args=(Path(sys.executable),),daemon=True).start()
            self._start_update_check()

    def _build(self):
        s=ttk.Style(self)
        s.theme_use("clam")
        s.configure(
            "Treeview",
            font=font(9),
            rowheight=30,
            background=CARD,
            fieldbackground=CARD,
            foreground=TEXT,
            borderwidth=0,
            bordercolor=CARD,
            lightcolor=CARD,
            darkcolor=CARD,
        )
        s.map("Treeview",background=[("selected",SELECT_BG)],foreground=[("selected",TEXT)])
        s.configure(
            "Treeview.Heading",
            font=font(9,True),
            padding=(6,8),
            background=CARD,
            foreground=MUTED,
            bordercolor=BORDER,
            lightcolor=CARD,
            darkcolor=CARD,
            relief="flat",
        )
        s.map("Treeview.Heading",background=[("active",CARD)])

        # One fixed-height layout: the result table takes all remaining height,
        # so the work area is never pushed below the fold or into a page scroll.
        # Grouped like a settings page: white rounded groups on a light grey ground,
        # hairlines between rows, no per-control borders.
        root=ctk.CTkFrame(self,fg_color=BG,corner_radius=0)
        root.pack(fill="both",expand=True,padx=16,pady=(14,12))
        self.root_content=root
        root.grid_columnconfigure(0,weight=1)
        root.grid_rowconfigure(4,weight=1)

        header=ctk.CTkFrame(root,fg_color="transparent")
        header.grid(row=0,column=0,sticky="ew",pady=(0,12))
        header.grid_columnconfigure(0,weight=1)
        ctk.CTkLabel(
            header,text="명세서 대사",font=font(24,True),text_color=TEXT,anchor="w"
        ).grid(row=0,column=0,sticky="w")
        ctk.CTkLabel(
            header,text="전월 말 명세서를 당월 더존 전표와 대사해 당월 말 명세서 초안을 만듭니다. 원본 파일은 수정하지 않습니다.",
            font=font(11),text_color=MUTED,anchor="w",justify="left",wraplength=380
        ).grid(row=1,column=0,sticky="w",pady=(3,0))
        self.update_btn=ctk.CTkButton(
            header,text="새 버전 업데이트",width=130,height=32,corner_radius=8,
            font=font(11,True),command=self._start_update,**BUTTON_PRIMARY
        )
        self.version_label=ctk.CTkLabel(
            header,text=(updater.version_label() if updater.version_label()=="개발 버전" else "버전 "+updater.version_label()) if updater.version_label() else "",
            font=font(10),text_color=MUTED,anchor="e"
        )
        self.version_label.grid(row=0,column=1,sticky="e")
        self.update_btn.grid(row=1,column=1,sticky="e")
        self.update_btn.grid_remove()

        def group(parent):
            return ctk.CTkFrame(parent,fg_color=CARD,corner_radius=12,border_width=0)

        def hairline(parent,row,columnspan=3):
            tk.Frame(parent,height=1,bg=BORDER,borderwidth=0).grid(
                row=row,column=0,columnspan=columnspan,sticky="ew",padx=(18,0)
            )

        # Inputs: period, two file lists, account code, in one group.
        form=group(root)
        self.form=form
        form.grid(row=1,column=0,sticky="ew")
        form.grid_columnconfigure(0,minsize=104)
        form.grid_columnconfigure(1,weight=1)
        form.grid_columnconfigure(2,minsize=150)

        def row_label(row,title,hint_var=None):
            box=ctk.CTkFrame(form,fg_color="transparent")
            box.grid(row=row,column=0,sticky="nw",padx=(14,6),pady=(10,0))
            ctk.CTkLabel(
                box,text=title,font=font(12,True),text_color=TEXT,anchor="w"
            ).pack(anchor="w")
            if hint_var is not None:
                ctk.CTkLabel(
                    box,textvariable=hint_var,font=font(10),text_color=MUTED,anchor="w",
                    justify="left",wraplength=92
                ).pack(anchor="w",pady=(2,0))

        row_label(0,"대상 회계월")
        period_box=ctk.CTkFrame(form,fg_color="transparent")
        period_box.grid(row=0,column=1,columnspan=2,sticky="ew",pady=(8,8),padx=(0,18))
        # One quiet line: the month, how it was chosen, and "변경". The pickers appear only
        # when detection failed or the files disagree, or when the user asks to change it.
        period_row=ctk.CTkFrame(period_box,fg_color="transparent")
        period_row.pack(fill="x")
        today=date.today()
        ctk.CTkLabel(
            period_row,textvariable=self.period_value_text,font=font(14,True),text_color=TEXT
        ).pack(side="left",padx=(0,12))
        # Plain text, not a badge: how the month was chosen.
        self.detect_pill=ctk.CTkLabel(period_row,text="",font=font(11,True),anchor="w")
        self.detect_pill.pack(side="left",padx=(0,6))
        self.period_toggle_btn=ctk.CTkButton(
            period_row,text="변경",command=self._toggle_period_picker,width=52,height=26,
            corner_radius=8,fg_color="transparent",hover_color=SOFT,text_color=ACCENT,
            font=font(12,True),border_width=0
        )
        self.period_toggle_btn.pack(side="left")
        ctk.CTkLabel(
            period_box,textvariable=self.period_text,font=font(10),text_color=MUTED,anchor="w"
        ).pack(anchor="w",pady=(2,0))
        self.period_picker=ctk.CTkFrame(period_box,fg_color="transparent")
        menu_style=dict(
            height=32,corner_radius=8,fg_color=SOFT,button_color=SOFT,
            button_hover_color=SOFT_HOVER,text_color=TEXT,dropdown_fg_color=CARD,
            dropdown_text_color=TEXT,dropdown_hover_color=SOFT,
            font=font(12,True),dropdown_font=font(12),
        )
        ctk.CTkOptionMenu(
            self.period_picker,values=[str(y) for y in range(today.year-10,today.year+2)],
            variable=self.year_text,command=self._on_period_menu,width=92,**menu_style
        ).pack(side="left")
        ctk.CTkLabel(self.period_picker,text="년",font=font(12),text_color=TEXT).pack(side="left",padx=(5,9))
        ctk.CTkOptionMenu(
            self.period_picker,values=[str(m) for m in range(1,13)],
            variable=self.month_text,command=self._on_period_menu,width=72,**menu_style
        ).pack(side="left")
        ctk.CTkLabel(self.period_picker,text="월",font=font(12),text_color=TEXT).pack(side="left",padx=(5,0))
        self._set_detect_pill("none")

        self._file_hints=[]

        def add_file_row(row,title,count_var,add_command,remove_selected,remove_all,placeholder=""):
            row_label(row,title,count_var)
            file_box=ctk.CTkFrame(
                form,fg_color=BG,corner_radius=8,border_width=0,height=42
            )
            file_box.grid(row=row,column=1,sticky="ew",pady=(8,8),padx=(0,2))
            file_box.grid_propagate(False)
            lb=tk.Listbox(
                file_box,height=2,font=font(10),selectmode="extended",
                bg=BG,fg=TEXT,selectbackground=SELECT_BG,selectforeground=TEXT,
                relief="flat",highlightthickness=0,borderwidth=0,
                activestyle="none",exportselection=False
            )
            lb.pack(side="left",fill="both",expand=True,padx=(12,8),pady=3)
            lb.bind("<Delete>",lambda _e:remove_selected())
            # Says what the empty box is for; shown only while the list has nothing in it.
            hint=ctk.CTkLabel(file_box,text=placeholder,font=font(11),text_color=MUTED,
                              fg_color=BG,anchor="w",cursor="hand2")
            hint.bind("<Button-1>",lambda _e:add_command())
            self._file_hints.append((lb,hint))

            actions=ctk.CTkFrame(form,fg_color="transparent")
            actions.grid(row=row,column=2,padx=(4,10),pady=(8,8),sticky="e")
            for i,(text,command,fg,bold) in enumerate((
                ("추가",add_command,ACCENT,True),
                ("선택 제거",remove_selected,TEXT2,False),
                ("모두 제거",remove_all,TEXT2,False),
            )):
                ctk.CTkButton(
                    actions,text=text,command=command,width=38 if i==0 else 56,height=28,
                    corner_radius=8,fg_color="transparent",hover_color=SOFT,
                    text_color=fg,font=font(11,bold),border_width=0
                ).pack(side="left",padx=(0 if i==0 else 1,0))
            return lb

        hairline(form,1)
        self.prior_list=add_file_row(
            2,"명세서",self.prior_count_text,self.pick_priors,
            self.remove_priors,self.clear_priors,
            "전월 말 명세서 파일을 추가하세요"
        )
        hairline(form,3)
        self.douzone_list=add_file_row(
            4,"더존 원장",self.raw_count_text,self.pick_douzone,
            self.remove_douzone,self.clear_douzone,
            "당월 더존 원장 파일을 추가하세요"
        )
        hairline(form,5)

        options=ctk.CTkFrame(form,fg_color="transparent")
        options.grid(row=6,column=0,columnspan=3,sticky="ew",padx=14,pady=(8,9))
        ctk.CTkLabel(
            options,text="계정코드",font=font(12,True),text_color=TEXT
        ).pack(side="left",padx=(0,10))
        ctk.CTkEntry(
            options,textvariable=self.account_codes,width=130,height=30,
            corner_radius=8,fg_color=BG,border_width=0,
            text_color=TEXT,font=font(12)
        ).pack(side="left")
        ctk.CTkLabel(
            options,text="미지급금 · 여러 개면 쉼표(,)로 구분",font=font(10),text_color=MUTED
        ).pack(side="left",padx=(10,0))

        # Actions and live status on one row.
        action_row=ctk.CTkFrame(root,fg_color="transparent")
        action_row.grid(row=2,column=0,sticky="ew",pady=(14,0))
        action_row.grid_columnconfigure(1,weight=1)
        buttons=ctk.CTkFrame(action_row,fg_color="transparent")
        buttons.grid(row=0,column=0,sticky="nw")
        self.run_btn=ctk.CTkButton(
            buttons,text="대사 시작",command=self.run,width=104,height=38,
            corner_radius=10,font=font(13,True),**BUTTON_PRIMARY
        )
        self.run_btn.pack(side="left")
        secondary=dict(height=38,corner_radius=10,font=font(12,True),border_width=0,**BUTTON_SECONDARY)
        self.month_end_btn=ctk.CTkButton(
            buttons,text="당월 명세서 생성",command=self.export_month_end,state="disabled",
            width=124,**secondary
        )
        self.month_end_btn.pack(side="left",padx=(8,0))
        # Compact input recap shown instead of the form once a run has finished.
        self.input_bar=group(root)
        self.input_bar.grid(row=1,column=0,sticky="ew")
        self.input_bar.grid_columnconfigure(0,weight=1)
        self.input_bar_text=ctk.CTkLabel(
            self.input_bar,text="",font=font(11),text_color=TEXT2,anchor="w"
        )
        self.input_bar_text.grid(row=0,column=0,sticky="ew",padx=(14,6),pady=10)
        ctk.CTkButton(
            self.input_bar,text="입력 수정",command=self._open_inputs,width=76,height=28,
            corner_radius=8,fg_color="transparent",hover_color=SOFT,text_color=ACCENT,
            font=font(11,True),border_width=0
        ).grid(row=0,column=1,padx=(0,8),pady=6)
        self.input_bar.grid_remove()

        # Status: plain text, no box. Colour carries only the state.
        self.banner=ctk.CTkFrame(action_row,fg_color="transparent")
        self.banner.grid(row=1,column=0,columnspan=2,sticky="ew",pady=(8,0))
        banner_inner=ctk.CTkFrame(self.banner,fg_color="transparent")
        banner_inner.pack(fill="x")
        banner_inner.grid_columnconfigure(0,weight=1)
        self.banner_title=ctk.CTkLabel(
            banner_inner,text="대사 전",font=font(13,True),text_color=TEXT,anchor="w"
        )
        self.banner_title.grid(row=0,column=0,sticky="w")
        self.banner_detail=ctk.CTkLabel(
            banner_inner,textvariable=self.status_text,font=font(11),
            text_color=MUTED,anchor="w",justify="left",wraplength=440
        )
        self.banner_detail.grid(row=1,column=0,sticky="w")
        self.progress=ctk.CTkProgressBar(
            banner_inner,mode="indeterminate",width=96,height=5,
            progress_color=ACCENT,fg_color=BORDER
        )
        self.progress.grid(row=0,column=1,rowspan=2,padx=(10,0))
        self.progress.grid_remove()
        self.banner.bind("<Configure>",self._on_banner_configure)

        # Summary: one line of plain numbers, colour only where it carries meaning.
        self.summary=ctk.CTkFrame(root,fg_color="transparent")
        self.summary.grid(row=3,column=0,sticky="ew",pady=(12,8))
        self._summary({})

        # Result: the main work area.
        result=ctk.CTkFrame(root,fg_color="transparent")
        result.grid(row=4,column=0,sticky="nsew")
        # Before a run there is nothing to list: show how the flow goes instead of an empty table.
        self.result_empty=group(result)
        ctk.CTkLabel(
            self.result_empty,text="이렇게 진행합니다",font=font(15,True),
            text_color=TEXT,anchor="w"
        ).pack(anchor="w",padx=18,pady=(16,4))
        for number,title,detail in (
            ("1","파일 추가","전월 말 명세서와 당월 더존 원장을 추가합니다. 회계월은 파일에서 자동으로 잡힙니다."),
            ("2","대사 시작","자동으로 짝을 맞추고, 사람이 봐야 할 항목만 추려 줍니다."),
            ("3","당월 명세서 생성","결과를 검토한 뒤 저장합니다. 전월 명세서 양식 그대로 나옵니다."),
        ):
            step=ctk.CTkFrame(self.result_empty,fg_color="transparent")
            step.pack(fill="x",padx=18,pady=(8,2))
            step.grid_columnconfigure(1,weight=1)
            badge=ctk.CTkFrame(step,width=26,height=26,corner_radius=13,fg_color=ACCENT)
            badge.grid(row=0,column=0,rowspan=2,sticky="n",padx=(0,12))
            badge.grid_propagate(False)
            ctk.CTkLabel(badge,text=number,font=font(11,True),text_color="white",
                         fg_color="transparent").place(relx=0.5,rely=0.5,anchor="center")
            ctk.CTkLabel(step,text=title,font=font(12,True),text_color=TEXT,anchor="w").grid(
                row=0,column=1,sticky="w")
            ctk.CTkLabel(step,text=detail,font=font(11),text_color=MUTED,anchor="w",
                         justify="left",wraplength=360).grid(
                row=1,column=1,sticky="w",pady=(1,0))
        ctk.CTkFrame(self.result_empty,height=14,fg_color="transparent").pack()
        result_main=ctk.CTkFrame(result,fg_color="transparent")
        self.result_main=result_main

        result_header=ctk.CTkFrame(result_main,fg_color="transparent")
        result_header.pack(fill="x",pady=(0,8))
        ctk.CTkLabel(
            result_header,text="결과 상세",font=font(15,True),text_color=TEXT
        ).pack(side="left")
        # Segmented control: a grey track with a white thumb on the active segment.
        segments=ctk.CTkFrame(result_header,fg_color=SOFT,corner_radius=9)
        segments.pack(side="right")
        view_style=dict(height=26,corner_radius=7,font=font(11,True),border_width=0)
        self.all_view_btn=ctk.CTkButton(
            segments,text="전체",command=lambda:self._set_result_view("all"),
            width=48,fg_color=CARD,hover_color=CARD,text_color=TEXT,**view_style
        )
        self.all_view_btn.pack(side="left",padx=(2,0),pady=2)
        self.review_view_btn=ctk.CTkButton(
            segments,text="검토 필요 0",command=lambda:self._set_result_view("review"),
            width=86,fg_color="transparent",hover_color=SOFT_HOVER,text_color=TEXT2,**view_style
        )
        self.review_view_btn.pack(side="left",padx=(0,0),pady=2)
        self.matched_view_btn=ctk.CTkButton(
            segments,text="자동 대사 0",command=lambda:self._set_result_view("matched"),
            width=86,fg_color="transparent",hover_color=SOFT_HOVER,text_color=TEXT2,**view_style
        )
        self.matched_view_btn.pack(side="left",padx=(0,2),pady=2)
        # Rarely needed, so it lives next to the results rather than among the main actions.
        self.save_btn=ctk.CTkButton(
            result_header,text="대사 결과 저장",command=self.export,state="disabled",
            width=84,height=26,corner_radius=8,fg_color="transparent",hover_color=SOFT,
            text_color=TEXT2,text_color_disabled=DISABLED_FG,font=font(11,True),border_width=0
        )
        self.save_btn.pack(side="right",padx=(0,6))

        hints=ctk.CTkFrame(result_main,fg_color="transparent")
        hints.pack(fill="x",pady=(0,6))
        ctk.CTkLabel(
            hints,textvariable=self.detail_hint_text,font=font(11),
            text_color=MUTED,anchor="w"
        ).pack(anchor="w")
        self.preflight=ctk.CTkLabel(
            hints,text="",font=font(11,True),text_color=WARN,anchor="w",
            justify="left",wraplength=440
        )
        self.preflight.pack(anchor="w")

        # The reason of the selected row, read in full (the table clips long text).
        detail_panel=group(result_main)
        detail_panel.pack(side="bottom",fill="x",pady=(8,0))
        panel_head=ctk.CTkFrame(detail_panel,fg_color="transparent")
        panel_head.pack(fill="x",padx=14,pady=(9,0))
        ctk.CTkLabel(
            panel_head,text="선택한 항목의 사유",font=font(12,True),text_color=TEXT2
        ).pack(side="left")
        ctk.CTkButton(
            panel_head,text="사유 복사",command=self._copy_detail,width=72,height=26,
            corner_radius=8,fg_color=SOFT,hover_color=SOFT_HOVER,text_color=TEXT2,
            font=font(11,True),border_width=0
        ).pack(side="right")
        self.detail_text=ctk.CTkTextbox(
            detail_panel,height=58,font=font(12),fg_color=CARD,text_color=TEXT,
            border_width=0,corner_radius=0,wrap="word"
        )
        self.detail_text.pack(fill="x",padx=10,pady=(2,8))
        self._show_detail("표에서 항목을 선택하면 사유 전체가 여기에 표시됩니다.")

        table_wrap=group(result_main)
        table_wrap.pack(fill="both",expand=True)
        table_wrap.grid_columnconfigure(0,weight=1)
        table_wrap.grid_rowconfigure(0,weight=1)
        self.detail=ttk.Treeview(
            table_wrap,
            columns=("owner","vendor","amount","status","reason"),
            show="headings",height=3,selectmode="browse",
            displaycolumns=("vendor","amount","status","reason"),
        )
        columns=(
            ("owner","원본 명세",130,False,"w"),
            ("vendor","거래처",118,False,"w"),
            ("amount","금액",80,False,"e"),
            ("status","상태",104,False,"w"),
            ("reason","사유 / 참고",160,True,"w"),
        )
        for col,title,width,stretch,anchor in columns:
            self.detail.heading(col,text=title,anchor=anchor)
            self.detail.column(col,width=width,minwidth=60,anchor=anchor,stretch=stretch)
        self.detail.tag_configure("matched",foreground=GREEN)
        self.detail.tag_configure("review",foreground=WARN)
        vsb=ctk.CTkScrollbar(
            table_wrap,orientation="vertical",command=self.detail.yview,width=10,
            fg_color=CARD,button_color="#C7C7CC",button_hover_color="#AEAEB2"
        )
        self.detail.configure(yscrollcommand=vsb.set)
        self.detail.grid(row=0,column=0,sticky="nsew",padx=(8,0),pady=8)
        vsb.grid(row=0,column=1,sticky="ns",padx=(0,5),pady=8)
        self.detail.bind("<<TreeviewSelect>>",self._on_detail_select)

        # Keyboard shortcuts for repeated monthly use.
        self.bind_all("<Control-Return>",lambda _e:self.run())
        self.bind_all("<Control-Key-1>",lambda _e:self.pick_priors())
        self.bind_all("<Control-Key-2>",lambda _e:self.pick_douzone())

        self.account_codes.trace_add("write",lambda *_:self._on_account_code_changed())
        self._inputs_open=False
        self._apply_state_layout()

    UPDATE_CHECK_MS=30*60*1000

    def _start_update_check(self):
        """Ask GitHub in the background (never blocks the window), then look again in 30 min."""
        threading.Thread(target=lambda:self._update_queue.put(updater.check_latest()),daemon=True).start()
        self.after(500,self._poll_update_check)

    def _poll_update_check(self):
        try:
            release=self._update_queue.get_nowait()
        except queue.Empty:
            self.after(500,self._poll_update_check)
            return
        if release is not None:
            self._update_release=release
            self.update_btn.grid()
        self.after(self.UPDATE_CHECK_MS,self._timer_update_check)

    def _timer_update_check(self):
        # Not while a run / file analysis / download is going on; try again next time.
        if self._run_busy or self._file_add_busy or str(self.update_btn.cget("state"))=="disabled":
            self.after(self.UPDATE_CHECK_MS,self._timer_update_check)
            return
        self._start_update_check()

    def _start_update(self):
        release=self._update_release
        if release is None or self._run_busy:
            return
        if not messagebox.askyesno(
            "업데이트",
            (f"현재 버전: {updater.version_label() or '알 수 없음'}\n"
             f"새 버전: {release.date or '최신'}\n\n")+
            "새 버전을 내려받아 프로그램을 다시 시작합니다.\n"
            "작업 중인 대사 결과는 사라지니, 필요하면 먼저 저장해 주세요.\n\n업데이트할까요?"
        ):
            return
        exe=Path(sys.executable)
        staged=updater.stage_path(exe)
        self.update_btn.configure(state="disabled",text="내려받는 중… 0%")
        updater.stop_old_swap_scripts()   # leftovers of an earlier failed update
        q=queue.Queue()

        def worker():
            try:
                # The button may have been up for a while: use the release as it is now.
                fresh=updater.fetch_release(force=True)
                target=fresh if fresh is not None and updater.is_newer(updater.current_sha(),fresh.sha) else release
                updater.download(target,staged,lambda done,total:q.put(("p",done,total)))
                q.put(("ok",))
            except Exception as e:
                try: staged.unlink()
                except OSError: pass
                q.put(("err",str(e) or "내려받기에 실패했습니다. 인터넷 연결을 확인해 주세요."))

        def poll():
            try:
                while True:
                    msg=q.get_nowait()
                    if msg[0]=="p":
                        pct=int(msg[1]*100/msg[2]) if msg[2] else 0
                        self.update_btn.configure(text=f"내려받는 중… {pct}%")
                    elif msg[0]=="err":
                        self.update_btn.configure(state="normal",text=self._update_text)
                        messagebox.showerror("업데이트 실패",msg[1]+"\n\n지금 쓰던 버전은 그대로입니다.")
                        return
                    else:
                        try:
                            updater.launch_finisher(exe,staged)
                        except OSError as e:
                            try: staged.unlink()
                            except OSError: pass
                            self.update_btn.configure(state="normal",text=self._update_text)
                            messagebox.showerror("업데이트 실패",f"교체를 시작하지 못했습니다: {e}")
                            return
                        self.destroy()
                        return
            except queue.Empty:
                pass
            self.after(100,poll)

        threading.Thread(target=worker,daemon=True).start()
        self.after(100,poll)

    def _sync_file_hints(self):
        try:
            for lb,hint in self._file_hints:
                if lb.size()==0:
                    hint.place(x=12,rely=0.5,anchor="w")
                else:
                    hint.place_forget()
            self.after(300,self._sync_file_hints)
        except tk.TclError:
            pass   # the window was closed

    def _style_button(self,btn,primary):
        btn.configure(**(BUTTON_PRIMARY if primary else BUTTON_SECONDARY))

    def _apply_state_layout(self):
        """Show what the next step needs: the primary button, the inputs, or the results."""
        if not hasattr(self,"result_main") or not hasattr(self,"_inputs_open"):
            return
        done=self.last_reconciliation_run is not None and not self._run_busy
        self._style_button(self.run_btn,not done)
        self._style_button(self.month_end_btn,done)
        if done:
            self.result_empty.pack_forget()
            self.result_main.pack(fill="both",expand=True)
            self.summary.grid()
        else:
            self.result_main.pack_forget()
            self.result_empty.pack(fill="x")
            self.summary.grid_remove()
        if done and not self._inputs_open:
            codes=", ".join(sorted(parse_account_codes(self.account_codes.get()))) or "-"
            period=self.last_period
            label=period.short_label if period else ""
            self.input_bar_text.configure(
                text=f"명세서 {len(self.prior_paths)}개  ·  더존 원장 {len(self.douzone_paths)}개"
                     f"  ·  계정코드 {codes}  ·  {label} 대상"
            )
            self.form.grid_remove()
            self.input_bar.grid()
        else:
            self.input_bar.grid_remove()
            self.form.grid()

    def _open_inputs(self):
        self._inputs_open=True
        self._apply_state_layout()

    def _on_banner_configure(self,event):
        # Keep the status text wrapped inside the banner on any window width.
        self.banner_detail.configure(wraplength=max(240,event.width-110))

    def _on_account_code_changed(self):
        self._invalidate_results()
        if self.prior_paths or self.douzone_paths:
            self.after_idle(lambda:self._maybe_align_period_from_inputs())

    def _set_banner(self,kind,title,detail):
        title_fg={"idle":TEXT,"running":TEXT,"success":GREEN,"warning":WARN,"error":RED}[kind]
        self.banner_title.configure(text=title,text_color=title_fg)
        self.banner_detail.configure(text_color=MUTED)
        self.status_text.set(detail)

    def _set_detect_pill(self,state):
        """Show how the accounting month was chosen, so it never claims more than is true."""
        text,fg={
            "none":("파일을 추가하면 자동 감지",MUTED),
            "auto":("파일에서 자동 설정됨",GREEN),
            "manual":("직접 선택함",TEXT2),
            "undetected":("월을 감지하지 못함 · 직접 선택",WARN),
            "conflict":("파일의 월과 다름",WARN),
        }[state]
        self.detect_pill.configure(text=text,text_color=fg)
        self._detect_state=state
        self._sync_period_picker()

    def _sync_period_picker(self):
        """The year/month pickers are a fallback: shown when detection failed or the files
        disagree (the user has to choose), or when the user pressed 변경."""
        if not hasattr(self,"period_picker"):
            return
        forced=self._detect_state in ("undetected","conflict")
        # Forced open only while the situation lasts: once detection succeeds again the pickers
        # fold back, unless the user opened them with 변경.
        if forced or self._period_picker_open:
            self.period_picker.pack(fill="x",pady=(8,0),anchor="w")
        else:
            self.period_picker.pack_forget()
        if forced:
            self.period_toggle_btn.pack_forget()
        else:
            self.period_toggle_btn.pack(side="left")
            self.period_toggle_btn.configure(text="접기" if self._period_picker_open else "변경")

    def _toggle_period_picker(self):
        self._period_picker_open=not self._period_picker_open
        self._sync_period_picker()

    def _update_period(self):
        try:
            p=AccountingPeriod(int(self.year.get()),int(self.month.get()))
            self.period_text.set(
                f"{p.previous().short_label} 말 명세서 + {p.short_label} 원장 → {p.short_label} 말 초안"
            )
            self.year_text.set(str(p.year))
            self.month_text.set(str(p.month))
            self.period_value_text.set(f"{p.year}년 {p.month}월")
        except Exception:
            self.period_text.set("파일을 추가하면 회계월을 자동으로 감지합니다.")
            self.period_value_text.set("-")
        if hasattr(self,"prior_count_text"):
            self._refresh_file_counts()

    def _on_period_menu(self,_value=None):
        """The user picked the month: keep it, and tell them if the files disagree."""
        try:
            self.year.set(int(self.year_text.get()))
            self.month.set(int(self.month_text.get()))
        except (ValueError,tk.TclError):
            return
        self._period_manual=True
        self._set_detect_pill("manual")
        self._on_period_changed()
        if self.prior_paths or self.douzone_paths:
            self.after_idle(lambda:self._maybe_align_period_from_inputs())

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
        self.save_btn.configure(state="disabled")
        self._inputs_open=False
        self._summary({})
        self.preflight.configure(text="")
        self._clear_detail()
        self.result_view="review"
        self._update_view_buttons()
        self.detail_hint_text.set("대사를 실행하면 검토가 필요한 항목을 여기에 보여줍니다.")
        if had_result:
            self._set_banner("warning","입력 조건 변경","입력 조건이 변경되었습니다. 다시 대사를 실행하세요.")
        else:
            self._set_banner("idle","대사 전",IDLE_HINT)
        self._apply_state_layout()

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
            self._set_detect_pill("manual" if self._period_manual else "none")
            self.period_text.set("명세서와 더존 원장 파일을 추가하면 회계월을 자동으로 감지합니다.")
            self._refresh_file_counts()
            if show_banner:
                self._set_banner("idle","대사 전",IDLE_HINT)
            return None

        statement_groups=[]
        statement_labels=[]
        for path in self.prior_paths:
            try:
                periods=(detected[path] if detected is not None and path in detected
                         else cached_periods(self._period_cache,"prior",path,()))
            except Exception:
                periods=[]
            if periods:
                statement_groups.append(periods)
                statement_labels.extend(periods)

        codes=parse_account_codes(self.account_codes.get())
        raw_groups=[]
        raw_labels=[]
        for path in self.douzone_paths:
            try:
                periods=(detected[path] if detected is not None and path in detected
                         else cached_periods(self._period_cache,"douzone",path,codes))
            except Exception:
                periods=[]
            if periods:
                raw_groups.append(periods)
                raw_labels.extend(periods)

        current=AccountingPeriod(int(self.year.get()),int(self.month.get()))
        suggested,source=infer_period_from_inputs(statement_groups,raw_groups,preferred=current)
        if suggested is None:
            if source=="conflict":
                self._set_detect_pill("conflict")
                if show_banner:
                    statement_text=", ".join(p.label for p in sorted(set(statement_labels))) or "감지 안 됨"
                    raw_text=", ".join(p.label for p in sorted(set(raw_labels))) or "감지 안 됨"
                    self._set_banner(
                        "warning",
                        "회계월 자동 설정 보류",
                        f"전월 명세서({statement_text})와 당월 더존 원장({raw_text})가 이어지지 않습니다."
                    )
                return False
            self._set_detect_pill("manual" if self._period_manual else "undetected")
            return None

        if suggested!=current:
            if self._period_manual:
                # Never overwrite a month the user chose; say how it differs instead.
                self._set_detect_pill("conflict")
                if show_banner:
                    self._set_banner(
                        "warning","선택한 회계월 확인",
                        f"선택한 {current.label}과 파일에서 감지한 {suggested.label}이 다릅니다. "
                        "월을 다시 선택하거나 파일을 확인해주세요."
                    )
                return False
            # Set both variables without leaving a stale result behind.
            self.year.set(suggested.year)
            self.month.set(suggested.month)
            self._update_period()
        self._set_detect_pill("manual" if self._period_manual else "auto")

        if show_banner:
            if source=="statement+raw":
                detail=f"{suggested.previous().label} 명세서 + {suggested.label} 더존 원장를 확인해 대상 회계월을 자동 설정했습니다."
            elif source=="statement":
                detail=f"전월 명세서 기준으로 당월 회계월을 {suggested.label}로 자동 설정했습니다."
            else:
                detail=f"더존 원장의 미지급금 전표월을 기준으로 대상 회계월을 {suggested.label}로 자동 설정했습니다."
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
        codes=parse_account_codes(self.account_codes.get())
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
                    try: detected[p]=cached_periods(self._period_cache,"prior",p,codes)
                    except Exception: detected[p]=[]
                for p in new_raw:
                    try: detected[p]=cached_periods(self._period_cache,"douzone",p,codes)
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
                        moved.append(f"{Path(p).name} → 더존 원장")
            elif kind=="ambiguous":
                rejected.append(f"{Path(p).name}: 명세서와 더존 구조가 함께 보여 자동 분류하지 않음")
            else:
                rejected.append(f"{Path(p).name}: 명세서/더존 원장 구조를 식별하지 못함")

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
            self._period_manual=False
            self._add_classified_files(paths,"prior")

    def pick_douzone(self):
        paths=filedialog.askopenfilenames(filetypes=[("Excel","*.xlsx *.xlsm *.xls")])
        if paths:
            self._period_manual=False
            self._add_classified_files(paths,"douzone")

    def _remove_selected(self,listbox,paths):
        """Remove only the files the user selected (a single file needs no selection)."""
        if self._file_add_busy or not paths:
            return
        picked=sorted(listbox.curselection(),reverse=True)
        if not picked:
            if len(paths)==1:
                picked=[0]
            else:
                self._set_banner(
                    "warning","제거할 파일 선택",
                    "목록에서 제거할 파일을 먼저 선택하세요. 전부 지우려면 '모두 제거'를 누르세요."
                )
                return
        for index in picked:
            if 0<=index<len(paths):
                listbox.delete(index)
                del paths[index]
        self._after_inputs_removed()

    def _clear_all(self,listbox,paths):
        if self._file_add_busy or not paths:
            return
        listbox.delete(0,"end")
        paths.clear()
        self._after_inputs_removed()

    def _after_inputs_removed(self):
        self._period_manual=False
        self._refresh_file_counts()
        self._invalidate_results()
        self._maybe_align_period_from_inputs()

    def remove_priors(self):
        self._remove_selected(self.prior_list,self.prior_paths)

    def remove_douzone(self):
        self._remove_selected(self.douzone_list,self.douzone_paths)

    def clear_priors(self):
        self._clear_all(self.prior_list,self.prior_paths)

    def clear_douzone(self):
        self._clear_all(self.douzone_list,self.douzone_paths)

    def _summary(self,counts):
        for w in self.summary.winfo_children():
            w.destroy()
        b=reconciliation_breakdown(counts,len(self.new_items),len(self.issues))
        # Colour only the numbers that ask for attention; a zero stays quiet.
        items=[
            ("자동 대사 완료",b["matched"],GREEN),
            ("검토 필요",b["review"],WARN),
            ("입력 확인",b["issues"],TEXT),
            ("당월 신규",b["new"],TEXT),
        ]
        for i,(name,value,color) in enumerate(items):
            ctk.CTkLabel(
                self.summary,text=name,font=font(11),text_color=TEXT2
            ).pack(side="left",padx=(0 if i==0 else 14,4))
            ctk.CTkLabel(
                self.summary,text=f"{value:,}건",font=font(12,True),
                text_color=color if value else MUTED
            ).pack(side="left")
        self._update_view_buttons()

    def _clear_detail(self):
        for row in self.detail.get_children():
            self.detail.delete(row)
        self._detail_reasons={}
        if hasattr(self,"detail_text"):
            self._show_detail("표에서 항목을 선택하면 사유 전체가 여기에 표시됩니다.")

    def _show_detail(self,text):
        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0","end")
        self.detail_text.insert("1.0",text)
        self.detail_text.configure(state="disabled")

    def _on_detail_select(self,_event=None):
        selected=self.detail.selection()
        if selected:
            self._show_detail(self._detail_reasons.get(selected[0],""))

    def _copy_detail(self):
        text=self.detail_text.get("1.0","end").strip()
        if text:
            self.clipboard_clear()
            self.clipboard_append(text)

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
                fg_color=CARD if active else "transparent",
                hover_color=CARD if active else SOFT_HOVER,
                text_color=TEXT if active else TEXT2,
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

        first=None
        for x in rows:
            matched=x.status in (Status.MATCHED,Status.SIGNED_NET_AUTO)
            amount=f"{int(x.closing_balance if x.status==Status.SIGNED_NET_AUTO else x.prior.amount):,}"
            owner=x.prior.source.owner or x.prior.source.file_name
            # A mark next to the colour, so status is never colour-only.
            status_text=("✓ " if matched else "! ")+x.status.value
            iid=self.detail.insert(
                "","end",
                values=(owner,x.prior.vendor_name,amount,status_text,x.reason),
                tags=("matched" if matched else "review",),
            )
            self._detail_reasons[iid]=(
                f"{x.prior.vendor_name} · {amount}원 · {x.status.value}\n{x.reason}"
                +(f"\n원본 명세: {owner}" if owner else "")
            )
            if first is None:
                first=iid
        if first is not None:
            self.detail.selection_set(first)
            self.detail.focus(first)

    def run(self):
        if self._file_add_busy:
            messagebox.showinfo("파일 분석 중","파일 분석이 완료된 후 대사를 실행해주세요.")
            return
        if self._run_busy:
            return
        if not self.prior_paths or not self.douzone_paths:
            messagebox.showwarning("파일 필요","명세서와 더존 원장을 각각 1개 이상 추가해주세요.")
            return
        self._invalidate_results()
        sync_state=self._maybe_align_period_from_inputs(show_banner=False)
        if sync_state is False:
            self._set_banner("error","대사 중단","선택한 회계월과 파일의 월이 맞지 않습니다.")
            messagebox.showerror(
                "회계월 확인",
                "전월 말 명세서와 당월 더존 원장이 이어지는 월인지, "
                "선택한 회계월이 맞는지 확인해주세요.\n파일의 월과 선택한 월을 확인해주세요."
            )
            return
        try:
            p=AccountingPeriod(int(self.year.get()),int(self.month.get()))
        except Exception as e:
            self._set_banner("error","대사 중단","회계월을 확인한 뒤 다시 실행해주세요.")
            messagebox.showerror("대사 중단",str(e))
            return
        codes=parse_account_codes(self.account_codes.get())
        prior=list(self.prior_paths)
        raw=list(self.douzone_paths)
        source_paths=prior+raw
        self._set_running(True)
        self._set_banner("running","대사 중","사전검사와 자동 대사를 진행하고 있습니다. 큰 파일은 시간이 걸릴 수 있습니다.")

        def worker():
            # Heavy work (hashing, reading, matching) stays off the UI thread.
            try:
                before_digests=snapshot_file_digests(source_paths)
                dup=find_duplicate_files(prior,raw,before_digests)
                if dup:
                    a,b=dup[0]
                    raise ValueError(f"동일한 파일 내용이 중복 추가되었습니다: {a[1]} / {b[1]}")
                run=run_reconciliation(prior,raw,codes,p)
                after_digests=snapshot_file_digests(source_paths)
                if before_digests != after_digests:
                    raise ValueError("대사 실행 중 입력 Excel이 변경되었습니다. 파일 저장이 끝난 뒤 다시 실행해주세요.")
                self._run_queue.put(("ok",(p,source_paths,run,after_digests)))
            except Exception as e:
                self._run_queue.put(("error",e))

        threading.Thread(target=worker,daemon=True).start()
        self.after(100,self._poll_run)

    def _set_running(self,busy):
        self._run_busy=busy
        self.run_btn.configure(state="disabled" if busy else "normal",text="대사 중…" if busy else "대사 시작")
        if busy:
            self.month_end_btn.configure(state="disabled")
            self.save_btn.configure(state="disabled")
            self.progress.grid()
            self.progress.start()
        else:
            self.progress.stop()
            self.progress.grid_remove()
        self._apply_state_layout()

    def _poll_run(self):
        try:
            kind,payload=self._run_queue.get_nowait()
        except queue.Empty:
            self.after(150,self._poll_run)
            return
        self._set_running(False)
        if kind=="error":
            self._set_banner("error","대사 중단","입력 내용을 확인한 뒤 다시 실행해주세요.")
            messagebox.showerror("대사 중단",str(payload))
            return
        try:
            self._finish_run(*payload)
        except Exception as e:
            self._set_banner("error","대사 중단","입력 내용을 확인한 뒤 다시 실행해주세요.")
            messagebox.showerror("대사 중단",str(e))

    def _finish_run(self,p,source_paths,run,after_digests):
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
        if run.unseen_account_codes:
            # A typo in one of several codes would otherwise silently drop that account.
            detail+=" · 이 달 전표가 없는 계정코드: "+", ".join(run.unseen_account_codes)
        if breakdown["review"] or breakdown["issues"] or run.unseen_account_codes:
            self._set_banner("success","대사 완료 · 확인 항목 있음",detail)
        else:
            self._set_banner("success","대사 완료",detail)
        self.month_end_btn.configure(state="normal")
        self.save_btn.configure(state="normal")
        self._inputs_open=False
        self._apply_state_layout()

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
            notices=write_month_end_statement(
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
            if notices:
                # Nothing is blocked, but nothing is hidden either.
                shown="\n".join("· "+text for text in notices[:12])
                more=f"\n· 외 {len(notices)-12}건 (검토필요 시트 참고)" if len(notices)>12 else ""
                messagebox.showinfo(
                    "당월 명세서 생성 완료 · 확인할 곳이 있습니다",
                    "명세서는 만들었습니다. 읽을 수 없었거나 그대로 옮길 수 없었던 부분입니다. "
                    "항목마다 어떻게 처리했는지 적혀 있습니다.\n\n"
                    +shown+more
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
            return
        except ValueError as e:
            # The month-end statement is always built from the original statement
            # format. If that format cannot be kept, say exactly why and do not
            # save a differently shaped file in its place.
            self._set_banner("error","당월 명세서 생성 불가","전월 양식을 유지할 수 없어 생성하지 않았습니다.")
            messagebox.showerror(
                "당월 명세서 생성 실패",
                f"{e}\n\n"
                "전월 양식을 그대로 유지할 수 없어 당월 명세서를 만들지 않았습니다. "
                "다른 형태의 파일로 대신 만들지 않습니다.\n"
                "위 내용을 고친 뒤 다시 시도해 주세요. 대사 결과만 필요하면 '대사 결과 저장'을 사용하세요."
            )
            return
        except Exception as e:
            messagebox.showerror("당월 명세서 생성 실패",str(e))
            return
        try:
            open_saved_folder(path)
        except OSError:
            messagebox.showwarning(
                "저장 완료 · 폴더 열기 실패",
                "당월 명세서는 정상적으로 저장했습니다.\n"
                "저장 폴더를 자동으로 열지 못했습니다. 아래 위치를 확인해 주세요.\n\n"
                f"저장 위치: {Path(path).resolve()}"
            )

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
    if len(sys.argv)>=4 and sys.argv[1]==updater.FINISH_FLAG:
        # Second half of a self-update: we are the freshly downloaded exe.
        ok,message=updater.finish_update(Path(sys.argv[2]),Path(sys.argv[3]))
        if not ok:
            root=tk.Tk(); root.withdraw()
            messagebox.showerror(
                "업데이트 실패",
                message+"\n\n지금까지 쓰던 프로그램은 그대로입니다.\n"
                f"새 파일은 {sys.argv[3]} 에 있습니다. 필요하면 이 파일 이름을 PayableRecon.exe로 바꿔 쓰세요."
            )
        sys.exit(0 if ok else 1)
    if "--tk" not in sys.argv:
        # New web screen first; the classic screen stays as the fallback when no WebView2 engine exists.
        from webui.launch import run as run_web
        if run_web(selftest="--selftest" in sys.argv):
            sys.exit(0)
        if "--selftest" in sys.argv:
            out=os.environ.get("PAYABLE_SELFTEST_OUT")
            if out:
                Path(out).write_text("FAIL: web screen could not start (fell back)",encoding="utf-8")
            sys.exit(1)
    App().mainloop()
