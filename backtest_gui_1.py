import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
import json
import os
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

import backtest


class BacktestGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("국내 ETF 자산배분 백테스트 v3.0 (고급형)")
        self.geometry("1250x850")

        self.df_master = backtest.df_master.copy()

        if 'name' not in self.df_master.columns and 'ETF명' in self.df_master.columns:
            self.df_master = self.df_master.rename(columns={'ETF명': 'name'})
        if 'name' not in self.df_master.columns:
            self.df_master['name'] = '이름없음'

        self.port_a = []  # 티커 리스트 (순서 중요)
        self.port_b = []

        self.group_var = tk.StringVar(value="전체")
        self.sub_var = tk.StringVar(value="전체")
        self.search_var = tk.StringVar(value="")

        self.weight_entries_a = {}
        self.weight_entries_b = {}

        self.start_var = tk.StringVar(value="2018-01-01")
        self.initial_var = tk.StringVar(value="300000000")
        self.rf_var = tk.StringVar(value="0")
        self.rebalance_var = tk.StringVar(value="Monthly")
        self.benchmark_var = tk.StringVar(value="069500.KS") # KODEX 200

        self.sort_column = None
        self.sort_reverse = False

        # 드래그 앤 드롭용 상태
        self.drag_data = {"port": None, "iid": None}

        self._build_ui()

        self.search_var.trace_add("write", lambda *args: self._refresh_name_list())

        self._refresh_sub_options()
        self._refresh_group_options()
        self._refresh_name_list()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        root = ttk.Frame(self, padding=16)
        root.grid(row=0, column=0, sticky="nsew")
        root.columnconfigure(0, weight=1)
        root.columnconfigure(1, weight=1)
        root.rowconfigure(0, weight=1)
        root.rowconfigure(1, weight=0)

        left = ttk.Frame(root)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(9, weight=1)

        right = ttk.Frame(root)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=0)
        right.rowconfigure(1, weight=1)

        # 0) 전략 프리셋 & 파일 관리 (추가)
        ttk.Label(left, text="0) 전략 프리셋 / 설정 파일").grid(row=0, column=0, sticky="w")
        preset_frame = ttk.Frame(left)
        preset_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        self.preset_cb = ttk.Combobox(preset_frame, values=["선택 안함"] + list(backtest.STRATEGY_PRESETS.keys()), state="readonly")
        self.preset_cb.set("선택 안함")
        self.preset_cb.pack(side="left", fill="x", expand=True)
        ttk.Button(preset_frame, text="A 적용", command=lambda: self._apply_preset("A")).pack(side="left", padx=2)
        ttk.Button(preset_frame, text="B 적용", command=lambda: self._apply_preset("B")).pack(side="left")

        file_btn_frame = ttk.Frame(left)
        file_btn_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        ttk.Button(file_btn_frame, text="설정 저장", command=self._save_config).pack(side="left", fill="x", expand=True, padx=(0, 2))
        ttk.Button(file_btn_frame, text="설정 불러오기", command=self._load_config).pack(side="left", fill="x", expand=True)

        # 1) sub 선택
        ttk.Label(left, text="1) sub 선택").grid(row=3, column=0, sticky="w")
        self.sub_cb = ttk.Combobox(left, textvariable=self.sub_var, state="readonly")
        self.sub_cb.grid(row=4, column=0, sticky="ew", pady=(4, 8))
        self.sub_cb.bind("<<ComboboxSelected>>", lambda e: self._on_sub_changed())

        # 2) 그룹 선택
        ttk.Label(left, text="2) 그룹 선택").grid(row=5, column=0, sticky="w")
        self.group_cb = ttk.Combobox(left, textvariable=self.group_var, state="readonly")
        self.group_cb.grid(row=6, column=0, sticky="ew", pady=(4, 8))
        self.group_cb.bind("<<ComboboxSelected>>", lambda e: self._on_group_changed())

        # 검색
        ttk.Label(left, text="검색 (티커 또는 종목명)").grid(row=7, column=0, sticky="w")
        ttk.Entry(left, textvariable=self.search_var).grid(row=8, column=0, sticky="ew", pady=(4, 8))

        # 표 제목
        ttk.Label(left, text="3) 종목 선택 (더블클릭 시 포트 A 추가)").grid(row=9, column=0, sticky="w", pady=(4, 4))

        names_frame = ttk.Frame(left)
        names_frame.grid(row=10, column=0, sticky="nsew", pady=(0, 4))
        names_frame.columnconfigure(0, weight=1)
        names_frame.rowconfigure(0, weight=1)

        self.names_tree = ttk.Treeview(
            names_frame,
            columns=("ticker", "name", "listed"),
            show="headings",
            selectmode="extended"
        )
        self.names_tree.heading("ticker", text="티커", command=lambda: self._sort_tree("ticker"))
        self.names_tree.heading("name", text="종목명", command=lambda: self._sort_tree("name"))
        self.names_tree.heading("listed", text="상장", command=lambda: self._sort_tree("listed"))

        self.names_tree.column("ticker", width=80, anchor="center")
        self.names_tree.column("name", width=260, anchor="w")
        self.names_tree.column("listed", width=100, anchor="center")

        self.names_tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(names_frame, orient="vertical", command=self.names_tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        self.names_tree.configure(yscrollcommand=vsb.set)

        self.names_tree.bind("<Double-1>", lambda e: self._add_selected_to_port("A"))

        btns = ttk.Frame(left)
        btns.grid(row=11, column=0, sticky="ew", pady=(10, 0))
        btns.columnconfigure(0, weight=1)
        btns.columnconfigure(1, weight=1)

        ttk.Button(btns, text="포트 A 추가", command=lambda: self._add_selected_to_port("A")).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(btns, text="포트 B 추가", command=lambda: self._add_selected_to_port("B")).grid(row=0, column=1, sticky="ew", padx=(6, 0))

        # ── 오른쪽 패널 ──
        top_settings = ttk.LabelFrame(right, text="백테스트 설정", padding=12)
        top_settings.grid(row=0, column=0, sticky="ew")
        top_settings.columnconfigure(1, weight=1)
        top_settings.columnconfigure(3, weight=1)

        ttk.Label(top_settings, text="시작일").grid(row=0, column=0, sticky="w")
        ttk.Entry(top_settings, textvariable=self.start_var).grid(row=0, column=1, sticky="ew", padx=(6, 12))
        ttk.Label(top_settings, text="초기금액").grid(row=0, column=2, sticky="w")
        ttk.Entry(top_settings, textvariable=self.initial_var).grid(row=0, column=3, sticky="ew", padx=(6, 0))

        ttk.Label(top_settings, text="무위험수익률").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(top_settings, textvariable=self.rf_var).grid(row=1, column=1, sticky="ew", padx=(6, 12), pady=(8, 0))

        # 리밸런싱 및 벤치마크 (추가)
        ttk.Label(top_settings, text="리밸런싱").grid(row=1, column=2, sticky="w", pady=(8, 0))
        ttk.Combobox(top_settings, textvariable=self.rebalance_var, values=["Monthly", "Quarterly", "Yearly", "None"], state="readonly").grid(row=1, column=3, sticky="ew", padx=(6, 0), pady=(8, 0))

        ttk.Label(top_settings, text="벤치마크").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(top_settings, textvariable=self.benchmark_var).grid(row=2, column=1, sticky="ew", padx=(6, 12), pady=(8, 0))

        ttk.Label(top_settings, text="포트 A/B 비중은 아래에서 % 단위로 입력 (합계 자동 정규화)").grid(row=3, column=0, columnspan=4, sticky="w", pady=(10, 0))

        ports_frame = ttk.Frame(right)
        ports_frame.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        ports_frame.columnconfigure(0, weight=1)
        ports_frame.columnconfigure(1, weight=1)
        ports_frame.rowconfigure(0, weight=1)

        # 포트 A - Treeview 표
        port_a_frame = ttk.LabelFrame(ports_frame, text="포트폴리오 A", padding=10)
        port_a_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        port_a_frame.columnconfigure(0, weight=1)
        port_a_frame.rowconfigure(0, weight=1)

        self.port_a_tree = ttk.Treeview(
            port_a_frame,
            columns=("ticker", "name", "listed", "weight"),
            show="headings",
            selectmode="browse"
        )
        self.port_a_tree.heading("ticker", text="티커")
        self.port_a_tree.heading("name", text="종목명")
        self.port_a_tree.heading("listed", text="상장")
        self.port_a_tree.heading("weight", text="비중(%)")

        self.port_a_tree.column("ticker", width=80, anchor="center")
        self.port_a_tree.column("name", width=220, anchor="w")
        self.port_a_tree.column("listed", width=100, anchor="center")
        self.port_a_tree.column("weight", width=90, anchor="center")

        self.port_a_tree.grid(row=0, column=0, sticky="nsew")

        a_scroll = ttk.Scrollbar(port_a_frame, orient="vertical", command=self.port_a_tree.yview)
        a_scroll.grid(row=0, column=1, sticky="ns")
        self.port_a_tree.configure(yscrollcommand=a_scroll.set)

        # 드래그 앤 드롭 이벤트
        self.port_a_tree.bind("<Button-1>", self._start_drag)
        self.port_a_tree.bind("<B1-Motion>", self._drag_motion)
        self.port_a_tree.bind("<ButtonRelease-1>", self._drop)

        a_btns = ttk.Frame(port_a_frame)
        a_btns.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(a_btns, text="선택 삭제", command=lambda: self._remove_selected_from_port("A")).pack(side="left")
        ttk.Button(a_btns, text="전체 지우기", command=lambda: self._clear_port("A")).pack(side="left", padx=6)
        ttk.Button(a_btns, text="A → B 복사", command=lambda: self._copy_port("A", "B")).pack(side="right")

        self.port_a_weights_inner = ttk.Frame(port_a_frame)
        self.port_a_weights_inner.grid(row=3, column=0, columnspan=2, sticky="nsew")
        
        sum_f_a = ttk.Frame(port_a_frame)
        sum_f_a.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        ttk.Button(sum_f_a, text="동일비중", command=lambda: self._apply_equal_weights("A")).pack(side="left")
        ttk.Button(sum_f_a, text="정규화", command=lambda: self._normalize_weights("A")).pack(side="left", padx=4)
        self.port_a_sum_label = ttk.Label(sum_f_a, text="합계: 0.0%")
        self.port_a_sum_label.pack(side="right")

        # 포트 B - Treeview 표 (A와 동일 구조)
        port_b_frame = ttk.LabelFrame(ports_frame, text="포트폴리오 B", padding=10)
        port_b_frame.grid(row=0, column=1, sticky="nsew")
        port_b_frame.columnconfigure(0, weight=1)
        port_b_frame.rowconfigure(0, weight=1)

        self.port_b_tree = ttk.Treeview(
            port_b_frame,
            columns=("ticker", "name", "listed", "weight"),
            show="headings",
            selectmode="browse"
        )
        self.port_b_tree.heading("ticker", text="티커")
        self.port_b_tree.heading("name", text="종목명")
        self.port_b_tree.heading("listed", text="상장")
        self.port_b_tree.heading("weight", text="비중(%)")

        self.port_b_tree.column("ticker", width=80, anchor="center")
        self.port_b_tree.column("name", width=220, anchor="w")
        self.port_b_tree.column("listed", width=100, anchor="center")
        self.port_b_tree.column("weight", width=90, anchor="center")

        self.port_b_tree.grid(row=0, column=0, sticky="nsew")

        b_scroll = ttk.Scrollbar(port_b_frame, orient="vertical", command=self.port_b_tree.yview)
        b_scroll.grid(row=0, column=1, sticky="ns")
        self.port_b_tree.configure(yscrollcommand=b_scroll.set)

        self.port_b_tree.bind("<Button-1>", self._start_drag)
        self.port_b_tree.bind("<B1-Motion>", self._drag_motion)
        self.port_b_tree.bind("<ButtonRelease-1>", self._drop)

        b_btns = ttk.Frame(port_b_frame)
        b_btns.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(b_btns, text="선택 삭제", command=lambda: self._remove_selected_from_port("B")).pack(side="left")
        ttk.Button(b_btns, text="전체 지우기", command=lambda: self._clear_port("B")).pack(side="left", padx=6)
        ttk.Button(b_btns, text="B → A 복사", command=lambda: self._copy_port("B", "A")).pack(side="right")

        self.port_b_weights_inner = ttk.Frame(port_b_frame)
        self.port_b_weights_inner.grid(row=3, column=0, columnspan=2, sticky="nsew")
        
        sum_f_b = ttk.Frame(port_b_frame)
        sum_f_b.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        ttk.Button(sum_f_b, text="동일비중", command=lambda: self._apply_equal_weights("B")).pack(side="left")
        ttk.Button(sum_f_b, text="정규화", command=lambda: self._normalize_weights("B")).pack(side="left", padx=4)
        self.port_b_sum_label = ttk.Label(sum_f_b, text="합계: 0.0%")
        self.port_b_sum_label.pack(side="right")

        # ── 하단 실행 버튼 ──
        bottom = ttk.Frame(root)
        bottom.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(20, 0))
        bottom.columnconfigure(0, weight=1)

        style = ttk.Style()
        style.configure("Big.TButton", font=("맑은 고딕", 16, "bold"))

        run_btn = ttk.Button(bottom, text="백테스트 실행 (Run Backtest)", command=self._run, style="Big.TButton")
        run_btn.grid(row=0, column=0, sticky="ew", padx=40, pady=15)

    # ────────────────────────────────────────────────
    # 드래그 앤 드롭 재정렬 (Treeview용)
    # ────────────────────────────────────────────────

    def _start_drag(self, event):
        widget = event.widget
        iid = widget.identify_row(event.y)
        if not iid:
            return

        if widget == self.port_a_tree:
            self.drag_data["port"] = "A"
        elif widget == self.port_b_tree:
            self.drag_data["port"] = "B"
        else:
            return

        self.drag_data["iid"] = iid
        widget.config(cursor="hand2")
        widget.selection_set(iid)

    def _drag_motion(self, event):
        if not self.drag_data["port"]:
            return
        widget = event.widget
        iid = widget.identify_row(event.y)
        if iid:
            widget.selection_set(iid)

    def _drop(self, event):
        if not self.drag_data["port"]:
            return

        widget = event.widget
        drop_iid = widget.identify_row(event.y)
        if not drop_iid:
            widget.config(cursor="")
            self.drag_data = {"port": None, "iid": None}
            return

        drag_iid = self.drag_data["iid"]
        if drag_iid == drop_iid:
            widget.config(cursor="")
            self.drag_data = {"port": None, "iid": None}
            return

        if self.drag_data["port"] == "A":
            target = self.port_a
            tree = self.port_a_tree
        else:
            target = self.port_b
            tree = self.port_b_tree

        current_order = [tree.item(i)['values'][0] for i in tree.get_children()]  # 티커
        drag_ticker = tree.item(drag_iid)['values'][0]
        drop_ticker = tree.item(drop_iid)['values'][0]

        drag_idx = current_order.index(drag_ticker)
        drop_idx = current_order.index(drop_ticker)

        item = target.pop(drag_idx)
        target.insert(drop_idx, item)

        self._refresh_port_trees()

        widget.config(cursor="")
        self.drag_data = {"port": None, "iid": None}

    def _refresh_port_trees(self):
        def get_info(ticker):
            row = self.df_master[self.df_master['ticker'] == str(ticker)]
            if row.empty:
                return ticker, "이름없음", ""
            r = row.iloc[0]
            return ticker, r.get('name', '이름없음'), r.get('listed', '')

        # 포트 A
        for item in self.port_a_tree.get_children():
            self.port_a_tree.delete(item)
        for t in self.port_a:
            ticker, name, listed = get_info(t)
            weight_var = self.weight_entries_a.get(t)
            weight = weight_var.get() if weight_var else "0.00"
            self.port_a_tree.insert("", "end", values=(ticker, name, listed, weight))

        # 포트 B
        for item in self.port_b_tree.get_children():
            self.port_b_tree.delete(item)
        for t in self.port_b:
            ticker, name, listed = get_info(t)
            weight_var = self.weight_entries_b.get(t)
            weight = weight_var.get() if weight_var else "0.00"
            self.port_b_tree.insert("", "end", values=(ticker, name, listed, weight))

        self._refresh_weight_entries("A")
        self._refresh_weight_entries("B")

    # ────────────────────────────────────────────────
    # 비중 및 포트폴리오 관리
    # ────────────────────────────────────────────────

    def _apply_preset(self, which):
        p_name = self.preset_cb.get()
        if p_name == "선택 안함": return
        p_data = backtest.STRATEGY_PRESETS[p_name]
        target = self.port_a if which=="A" else self.port_b
        target.clear()
        for ticker in p_data.keys():
            target.append(ticker)
        self._refresh_port_trees()
        entries = self.weight_entries_a if which=="A" else self.weight_entries_b
        for t, w in p_data.items(): 
            if t in entries: entries[t].set(f"{w*100:.1f}")
        self._update_weight_sum(which)

    def _save_config(self):
        config = {
            "start": self.start_var.get(),
            "initial": self.initial_var.get(),
            "rf": self.rf_var.get(),
            "rebalance": self.rebalance_var.get(),
            "benchmark": self.benchmark_var.get(),
            "port_a": {t: v.get() for t, v in self.weight_entries_a.items()},
            "port_b": {t: v.get() for t, v in self.weight_entries_b.items()}
        }
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if path:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            messagebox.showinfo("저장 완료", "설정이 저장되었습니다.")

    def _load_config(self):
        path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if not path: return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
            self.start_var.set(cfg.get("start", "2018-01-01"))
            self.initial_var.set(cfg.get("initial", "300000000"))
            self.rf_var.set(cfg.get("rf", "0"))
            self.rebalance_var.set(cfg.get("rebalance", "Monthly"))
            self.benchmark_var.set(cfg.get("benchmark", "069500.KS"))
            self.port_a = list(cfg.get("port_a", {}).keys())
            self.port_b = list(cfg.get("port_b", {}).keys())
            self._refresh_port_trees()
            for t, w in cfg.get("port_a", {}).items():
                if t in self.weight_entries_a: self.weight_entries_a[t].set(w)
            for t, w in cfg.get("port_b", {}).items():
                if t in self.weight_entries_b: self.weight_entries_b[t].set(w)
            self._update_weight_sum("A"); self._update_weight_sum("B")
        except Exception as e:
            messagebox.showerror("오류", str(e))

    def _normalize_weights(self, which):
        ents = self.weight_entries_a if which=="A" else self.weight_entries_b
        if not ents: return
        try:
            res = {t: float(v.get() or 0) for t, v in ents.items()}
            s = sum(res.values())
            if s <= 0: return
            for t, v in ents.items():
                v.set(f"{(float(v.get() or 0) / s * 100):.2f}")
        except: pass

    def _refresh_sub_options(self):
        subs = sorted(self.df_master['sub'].dropna().unique().tolist())
        self.sub_cb["values"] = ["전체"] + subs
        if self.sub_var.get() not in self.sub_cb["values"]:
            self.sub_var.set("전체")

    def _refresh_group_options(self):
        s = self.sub_var.get()
        df = self.df_master
        if s != "전체":
            df = df[df['sub'] == s]
        groups = sorted(df['group'].dropna().unique().tolist())
        self.group_cb["values"] = ["전체"] + groups
        if self.group_var.get() not in self.group_cb["values"]:
            self.group_var.set("전체")

    def _filtered_df(self):
        s = self.sub_var.get()
        g = self.group_var.get()
        df = self.df_master.copy()
        if s != "전체":
            df = df[df['sub'] == s]
        if g != "전체":
            df = df[df['group'] == g]
        return df

    def _on_sub_changed(self):
        self.group_var.set("전체")
        self._refresh_group_options()
        self._refresh_name_list()

    def _on_group_changed(self):
        self._refresh_name_list()

    def _refresh_name_list(self):
        df = self._filtered_df()
        search_text = self.search_var.get().strip().lower()
        if search_text:
            mask = (df['ticker'].astype(str).str.lower().str.contains(search_text, na=False) |
                    df['name'].astype(str).str.lower().str.contains(search_text, na=False))
            df = df[mask]

        for item in self.names_tree.get_children():
            self.names_tree.delete(item)
        for _, row in df.iterrows():
            ticker = str(row.get("ticker", ""))
            name   = row.get("name", "이름없음")
            listed = str(row.get("listed", "")) if pd.notna(row.get("listed")) else ""
            self.names_tree.insert("", "end", values=(ticker, name, listed))

        if self.sort_column:
            self._sort_tree(self.sort_column, apply_only=True)

    def _sort_tree(self, col, apply_only=False):
        if not apply_only:
            self.sort_reverse = not self.sort_reverse if self.sort_column == col else False
            self.sort_column = col
        items = [(self.names_tree.set(k, col), k) for k in self.names_tree.get_children("")]
        def safe_key(val):
            try: return (0, float(val))
            except: return (1, str(val).lower())
        items.sort(key=lambda x: safe_key(x[0]), reverse=self.sort_reverse)
        for index, (val, k) in enumerate(items):
            self.names_tree.move(k, '', index)

    def _add_selected_to_port(self, which):
        selected = self.names_tree.selection()
        target = self.port_a if which == "A" else self.port_b
        added = 0
        for iid in selected:
            t = self.names_tree.item(iid, "values")[0]
            if t not in target:
                target.append(t)
                added += 1
        if added > 0: self._refresh_port_trees()

    def _remove_selected_from_port(self, which):
        tree = self.port_a_tree if which == "A" else self.port_b_tree
        sel = tree.selection()
        if not sel: return
        ticker = tree.item(sel[0])['values'][0]
        target = self.port_a if which == "A" else self.port_b
        if ticker in target: target.remove(ticker)
        self._refresh_port_trees()

    def _clear_port(self, which):
        if which == "A": self.port_a.clear()
        else: self.port_b.clear()
        self._refresh_port_trees()

    def _copy_port(self, src, dst):
        s_list, d_list = (self.port_a, self.port_b) if src == "A" else (self.port_b, self.port_a)
        d_list[:] = list(s_list)
        s_dict, d_dict = (self.weight_entries_a, self.weight_entries_b) if src=="A" else (self.weight_entries_b, self.weight_entries_a)
        copied = {k: v.get() for k, v in s_dict.items()}
        for k, v in copied.items():
            if k in d_dict: d_dict[k].set(v)
        self._refresh_port_trees()

    def _refresh_weight_entries(self, which):
        tickers = self.port_a if which == "A" else self.port_b
        inner = self.port_a_weights_inner if which == "A" else self.port_b_weights_inner
        entries_dict = self.weight_entries_a if which == "A" else self.weight_entries_b
        for child in inner.winfo_children(): child.destroy()
        entries_dict.clear()
        eq_val = f"{100 / (len(tickers) or 1):.2f}"
        for i, t in enumerate(tickers):
            row = self.df_master[self.df_master["ticker"] == str(t)]
            name = row.iloc[0]["name"] if not row.empty else t
            ttk.Label(inner, text=f"{t} {name[:12]}").grid(row=i, column=0, sticky="w", padx=(0, 4))
            var = tk.StringVar(value=eq_val)
            ttk.Entry(inner, textvariable=var, width=8).grid(row=i, column=1, sticky="w")
            entries_dict[t] = var
            var.trace_add("write", lambda *args, w=which: self._update_weight_sum(w))
        self._update_weight_sum(which)

    def _update_weight_sum(self, which):
        ents = self.weight_entries_a if which=="A" else self.weight_entries_b
        total = sum(float(v.get() or 0) for v in ents.values())
        lbl = self.port_a_sum_label if which=="A" else self.port_b_sum_label
        lbl.config(text=f"합계: {total:.1f}%")

    def _apply_equal_weights(self, which):
        ents = self.weight_entries_a if which=="A" else self.weight_entries_b
        if not ents: return
        eq = f"{100/len(ents):.2f}"
        for v in ents.values(): v.set(eq)

    def _get_direct_weights(self, which):
        ents = self.weight_entries_a if which=="A" else self.weight_entries_b
        res = {t: float(v.get() or 0) for t, v in ents.items()}
        s = sum(res.values())
        if s <= 0: raise ValueError(f"포트 {which} 비중 오류")
        return {t: v/s for t, v in res.items()}

    def _run(self):
        try:
            if not self.port_a or not self.port_b: raise ValueError("종목을 추가하세요.")
            w_a, w_b = self._get_direct_weights("A"), self._get_direct_weights("B")
            result = backtest.run_backtest(
                backtest.build_df_weights(self.df_master, w_a, w_b),
                start=self.start_var.get(), initial_investment=int(self.initial_var.get()),
                rf=float(self.rf_var.get()), rebalance=self.rebalance_var.get(),
                benchmark_ticker=self.benchmark_var.get(), show_plots=False, verbose=False
            )
            self._open_performance_and_graph_window(result)
        except Exception as e: messagebox.showerror("오류", str(e))

    def _open_performance_and_graph_window(self, result):
        win = tk.Toplevel(self); win.title("백테스트 결과 리포트"); win.geometry("1100x850")
        nb = ttk.Notebook(win); nb.pack(fill="both", expand=True, padx=10, pady=10)

        # 탭 1: 성과 비교
        t1 = ttk.Frame(nb); nb.add(t1, text="성과 비교")
        tree = ttk.Treeview(t1, columns=("지표", "Port A", "Port B", "Benchmark"), show="headings")
        for c in tree["columns"]: tree.heading(c, text=c); tree.column(c, width=120, anchor="center")
        tree.pack(fill="both", expand=True)
        for idx, row in result["metrics_compare"].iterrows():
            tree.insert("", "end", values=(idx, row.get("Port A", "-"), row.get("Port B", "-"), row.get("Benchmark", "-")))
        ttk.Label(t1, text=f"Port A ↔ Port B 상관계수 : {result.get('correlation_ab', 0):.3f}", font=("맑은 고딕", 12, "bold")).pack(pady=10)

        # 탭 2: 월별 수익률 (v2 추가)
        t2 = ttk.Frame(nb); nb.add(t2, text="월별 수익률(A)")
        m_tree = ttk.Treeview(t2, columns=["Year"] + [f"{m}월" for m in range(1,13)], show="headings")
        m_tree.heading("Year", text="연도"); m_tree.column("Year", width=60)
        for m in range(1,13): m_tree.heading(f"{m}월", text=f"{m}월"); m_tree.column(f"{m}월", width=65, anchor="center")
        m_tree.pack(fill="both", expand=True)
        matrix = result['monthly_matrix_a']
        for year in matrix.index:
            vals = [year] + [f"{v:.1f}%" if pd.notna(v) else "-" for v in matrix.loc[year]]
            m_tree.insert("", "end", values=vals)

        # 탭 3: 그래프 (v2 롤링 수익률 포함)
        t3 = ttk.Frame(nb); nb.add(t3, text="분석 그래프")
        fig = Figure(figsize=(10, 10)); axs = fig.subplots(3, 1)
        axs[0].plot(result["asset_values_a"], label="Port A"); axs[0].plot(result["asset_values_b"], label="Port B")
        axs[0].plot(result["asset_values_bench"], label="Benchmark", ls="--", color="gray"); axs[0].set_title("누적 자산 가치"); axs[0].legend(); axs[0].grid(True, ls=":")
        axs[1].plot(result['rolling_12m_a'], label="Port A 1Y Rolling"); axs[1].plot(result['rolling_12m_b'], label="Port B 1Y Rolling")
        axs[1].axhline(0, color="black", lw=1); axs[1].set_title("12개월 롤링 수익률 (%)"); axs[1].legend(); axs[1].grid(True, ls=":")
        axs[2].fill_between(result["drawdown_a"].index, result["drawdown_a"], color="blue", alpha=0.1); axs[2].plot(result["drawdown_a"], color="blue")
        axs[2].plot(result["drawdown_b"], color="orange"); axs[2].set_title("Drawdown (%)"); axs[2].grid(True, ls=":")
        fig.tight_layout(); canvas = FigureCanvasTkAgg(fig, master=t3); canvas.draw(); canvas.get_tk_widget().pack(fill="both", expand=True)

if __name__ == "__main__": BacktestGUI().mainloop()
