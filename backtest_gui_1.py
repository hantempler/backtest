import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
import json
import os
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

import backtest

# --- Color Constants ---
COLOR_BG = "#F8F9FA"  # Light grayish background
COLOR_CARD = "#FFFFFF" # White for cards
COLOR_PRIMARY = "#2C3E50" # Dark slate for text/headers
COLOR_ACCENT = "#3498DB"  # Modern blue for buttons
COLOR_SUCCESS = "#27AE60" # Emerald green
COLOR_BORDER = "#DEE2E6" # Soft gray for borders
FONT_MAIN = ("Segoe UI", 10)
FONT_BOLD = ("Segoe UI", 10, "bold")
FONT_TITLE = ("Segoe UI", 12, "bold")

class BacktestGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("국내 ETF 자산배분 백테스트 v3.0 (Modern UI)")
        self.geometry("1300x900")
        self.configure(bg=COLOR_BG)

        self.df_master = backtest.df_master.copy()
        if 'name' not in self.df_master.columns and 'ETF명' in self.df_master.columns:
            self.df_master = self.df_master.rename(columns={'ETF명': 'name'})
        if 'name' not in self.df_master.columns:
            self.df_master['name'] = '이름없음'

        self.port_a = []
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
        # 벤치마크 선택을 위한 매핑 정보
        self.benchmark_options = {"KOSPI 200": "069500.KS", "S&P 500": "SPY"}
        self.benchmark_display_var = tk.StringVar(value="KOSPI 200")
        
        self.sort_column = None
        self.sort_reverse = False
        self.drag_data = {"port": None, "iid": None}
        
        self._apply_styles()
        self._build_ui()

        self.search_var.trace_add("write", lambda *args: self._refresh_name_list())
        self._refresh_sub_options()
        self._refresh_group_options()
        self._refresh_name_list()

    def _apply_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        # General
        style.configure(".", font=FONT_MAIN, background=COLOR_BG, foreground=COLOR_PRIMARY)
        
        # Cards (Frames)
        style.configure("Card.TFrame", background=COLOR_CARD, relief="flat")
        
        # Labels
        style.configure("TLabel", background=COLOR_BG, foreground=COLOR_PRIMARY)
        style.configure("Header.TLabel", font=FONT_TITLE, background=COLOR_CARD)
        style.configure("CardLabel.TLabel", background=COLOR_CARD, font=FONT_BOLD)
        
        # Buttons
        style.configure("Accent.TButton", font=FONT_BOLD, foreground="white", background=COLOR_ACCENT, borderwidth=0)
        style.map("Accent.TButton", background=[("active", "#2980B9")])
        
        style.configure("Success.TButton", font=FONT_BOLD, foreground="white", background=COLOR_SUCCESS, borderwidth=0)
        style.map("Success.TButton", background=[("active", "#219150")])

        style.configure("Normal.TButton", font=FONT_MAIN)
        
        # Treeview
        style.configure("Treeview", font=FONT_MAIN, rowheight=30, background="white", fieldbackground="white")
        style.configure("Treeview.Heading", font=FONT_BOLD, background=COLOR_BORDER, relief="flat")
        style.map("Treeview", background=[("selected", COLOR_ACCENT)], foreground=[("selected", "white")])

        # Entry & Combobox
        style.configure("TEntry", padding=5)
        style.configure("TCombobox", padding=5)

    def _build_ui(self):
        root = ttk.Frame(self, padding=20)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=4) # Left sidebar
        root.columnconfigure(1, weight=6) # Right main
        root.rowconfigure(0, weight=1)

        # --- LEFT PANEL ---
        left_panel = ttk.Frame(root)
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 15))
        left_panel.columnconfigure(0, weight=1)
        
        # Card 1: Presets & Files
        c1 = ttk.Frame(left_panel, style="Card.TFrame", padding=15)
        c1.pack(fill="x", pady=(0, 15))
        ttk.Label(c1, text="전략 프리셋 & 파일", style="Header.TLabel").pack(anchor="w", pady=(0, 10))
        
        preset_f = ttk.Frame(c1, style="Card.TFrame")
        preset_f.pack(fill="x")
        self.preset_cb = ttk.Combobox(preset_f, values=["선택 안함"] + list(backtest.STRATEGY_PRESETS.keys()), state="readonly")
        self.preset_cb.set("선택 안함")
        self.preset_cb.pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(preset_f, text="A 적용", command=lambda: self._apply_preset("A"), width=6).pack(side="left", padx=2)
        ttk.Button(preset_f, text="B 적용", command=lambda: self._apply_preset("B"), width=6).pack(side="left")

        file_f = ttk.Frame(c1, style="Card.TFrame")
        file_f.pack(fill="x", pady=(10, 0))
        ttk.Button(file_f, text="설정 저장", command=self._save_config).pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(file_f, text="설정 불러오기", command=self._load_config).pack(side="left", fill="x", expand=True)

        # Card 2: Filter & Search
        c2 = ttk.Frame(left_panel, style="Card.TFrame", padding=15)
        c2.pack(fill="both", expand=True)
        ttk.Label(c2, text="종목 탐색", style="Header.TLabel").pack(anchor="w", pady=(0, 10))
        
        ttk.Label(c2, text="시장 분류", style="CardLabel.TLabel").pack(anchor="w")
        self.sub_cb = ttk.Combobox(c2, textvariable=self.sub_var, state="readonly")
        self.sub_cb.pack(fill="x", pady=(2, 10))
        self.sub_cb.bind("<<ComboboxSelected>>", lambda e: self._on_sub_changed())

        ttk.Label(c2, text="자산 그룹", style="CardLabel.TLabel").pack(anchor="w")
        self.group_cb = ttk.Combobox(c2, textvariable=self.group_var, state="readonly")
        self.group_cb.pack(fill="x", pady=(2, 10))
        self.group_cb.bind("<<ComboboxSelected>>", lambda e: self._on_group_changed())

        ttk.Label(c2, text="검색 (티커/종목명)", style="CardLabel.TLabel").pack(anchor="w")
        ttk.Entry(c2, textvariable=self.search_var).pack(fill="x", pady=(2, 15))

        # 종목 리스트 Treeview
        tree_f = ttk.Frame(c2, style="Card.TFrame")
        tree_f.pack(fill="both", expand=True)
        tree_f.columnconfigure(0, weight=1)
        tree_f.rowconfigure(0, weight=1)

        self.names_tree = ttk.Treeview(tree_f, columns=("ticker", "name", "listed"), show="headings", selectmode="extended")
        self.names_tree.heading("ticker", text="티커", command=lambda: self._sort_tree("ticker"))
        self.names_tree.heading("name", text="종목명", command=lambda: self._sort_tree("name"))
        self.names_tree.heading("listed", text="상장일", command=lambda: self._sort_tree("listed"))
        self.names_tree.column("ticker", width=70, anchor="center")
        self.names_tree.column("name", width=200, anchor="w")
        self.names_tree.column("listed", width=90, anchor="center")
        self.names_tree.grid(row=0, column=0, sticky="nsew")
        
        vsb = ttk.Scrollbar(tree_f, orient="vertical", command=self.names_tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        self.names_tree.configure(yscrollcommand=vsb.set)
        self.names_tree.bind("<Double-1>", lambda e: self._add_selected_to_port("A"))

        # Add buttons
        add_f = ttk.Frame(c2, style="Card.TFrame")
        add_f.pack(fill="x", pady=(10, 0))
        ttk.Button(add_f, text="+ 포트 A 추가", command=lambda: self._add_selected_to_port("A"), style="Accent.TButton").pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(add_f, text="+ 포트 B 추가", command=lambda: self._add_selected_to_port("B"), style="Accent.TButton").pack(side="left", fill="x", expand=True)

        # --- RIGHT PANEL ---
        right_panel = ttk.Frame(root)
        right_panel.grid(row=0, column=1, sticky="nsew")
        right_panel.columnconfigure(0, weight=1)
        right_panel.rowconfigure(1, weight=1)

        # Card 3: Backtest Settings
        c3 = ttk.Frame(right_panel, style="Card.TFrame", padding=15)
        c3.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        ttk.Label(c3, text="백테스트 환경 설정", style="Header.TLabel").grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))
        
        labels = ["시작일", "초기금액", "무위험수익률", "리밸런싱", "벤치마크"]
        vars = [self.start_var, self.initial_var, self.rf_var, self.rebalance_var, self.benchmark_display_var]
        
        for i, (l, v) in enumerate(zip(labels, vars)):
            r, c = divmod(i, 3)
            ttk.Label(c3, text=l, style="CardLabel.TLabel").grid(row=r+1, column=c*2, sticky="w", padx=(10 if c>0 else 0, 5), pady=5)
            if l == "리밸런싱":
                ttk.Combobox(c3, textvariable=v, values=["Monthly", "Quarterly", "Yearly", "None"], state="readonly", width=12).grid(row=r+1, column=c*2+1, sticky="ew")
            elif l == "벤치마크":
                ttk.Combobox(c3, textvariable=v, values=list(self.benchmark_options.keys()), state="readonly", width=12).grid(row=r+1, column=c*2+1, sticky="ew")
            else:
                ttk.Entry(c3, textvariable=v, width=15).grid(row=r+1, column=c*2+1, sticky="ew")

        # Card 4: Portfolios
        c4 = ttk.Frame(right_panel, style="Card.TFrame", padding=15)
        c4.grid(row=1, column=0, sticky="nsew")
        c4.columnconfigure(0, weight=1)
        c4.columnconfigure(1, weight=1)
        c4.rowconfigure(0, weight=1)

        def build_port_ui(parent, title, which):
            frame = ttk.Frame(parent, style="Card.TFrame")
            frame.columnconfigure(0, weight=1)
            frame.rowconfigure(1, weight=1)
            
            hdr = ttk.Frame(frame, style="Card.TFrame")
            hdr.grid(row=0, column=0, sticky="ew", pady=(0, 5))
            ttk.Label(hdr, text=title, style="Header.TLabel").pack(side="left")
            
            # Treeview for port
            tree = ttk.Treeview(frame, columns=("ticker", "name", "weight"), show="headings", height=8)
            tree.heading("ticker", text="티커")
            tree.heading("name", text="종목명")
            tree.heading("weight", text="비중(%)")
            tree.column("ticker", width=70, anchor="center")
            tree.column("name", width=180, anchor="w")
            tree.column("weight", width=70, anchor="center")
            tree.grid(row=1, column=0, sticky="nsew")
            
            # Drag events
            tree.bind("<Button-1>", self._start_drag)
            tree.bind("<B1-Motion>", self._drag_motion)
            tree.bind("<ButtonRelease-1>", self._drop)

            # Scrollbar
            sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
            sb.grid(row=1, column=1, sticky="ns")
            tree.configure(yscrollcommand=sb.set)

            # Weight Inputs Area
            win = ttk.Frame(frame, style="Card.TFrame", padding=(0, 10))
            win.grid(row=2, column=0, columnspan=2, sticky="nsew")
            
            # Bottom bar (Buttons & Sum)
            bb = ttk.Frame(frame, style="Card.TFrame")
            bb.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(5, 0))
            ttk.Button(bb, text="삭제", command=lambda: self._remove_selected_from_port(which), width=5).pack(side="left")
            ttk.Button(bb, text="동일비중", command=lambda: self._apply_equal_weights(which), width=8).pack(side="left", padx=5)
            ttk.Button(bb, text="정규화", command=lambda: self._normalize_weights(which), width=6).pack(side="left")
            
            sum_lbl = ttk.Label(bb, text="합계: 0.0%", style="CardLabel.TLabel")
            sum_lbl.pack(side="right")
            
            if which == "A":
                self.port_a_tree, self.port_a_weights_inner, self.port_a_sum_label = tree, win, sum_lbl
            else:
                self.port_b_tree, self.port_b_weights_inner, self.port_b_sum_label = tree, win, sum_lbl
            
            return frame

        build_port_ui(c4, "Portfolio A", "A").grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        build_port_ui(c4, "Portfolio B", "B").grid(row=0, column=1, sticky="nsew")

        # --- EXECUTE BUTTON ---
        btn_run = ttk.Button(right_panel, text="RUN BACKTEST", style="Success.TButton", command=self._run)
        btn_run.grid(row=2, column=0, sticky="ew", pady=(15, 0), ipady=10)

    # --- Methods (Logic remains similar but with UI updates) ---

    def _start_drag(self, event):
        widget = event.widget
        iid = widget.identify_row(event.y)
        if not iid: return
        self.drag_data["port"] = "A" if widget == self.port_a_tree else "B"
        self.drag_data["iid"] = iid
        widget.config(cursor="hand2")
        widget.selection_set(iid)

    def _drag_motion(self, event):
        if not self.drag_data["port"]: return
        iid = event.widget.identify_row(event.y)
        if iid: event.widget.selection_set(iid)

    def _drop(self, event):
        if not self.drag_data["port"]: return
        widget = event.widget
        drop_iid = widget.identify_row(event.y)
        if not drop_iid or self.drag_data["iid"] == drop_iid:
            widget.config(cursor=""); self.drag_data = {"port": None, "iid": None}; return

        target = self.port_a if self.drag_data["port"] == "A" else self.port_b
        current_order = [widget.item(i)['values'][0] for i in widget.get_children()]
        drag_ticker = widget.item(self.drag_data["iid"])['values'][0]
        drop_ticker = widget.item(drop_iid)['values'][0]

        drag_idx, drop_idx = current_order.index(drag_ticker), current_order.index(drop_ticker)
        item = target.pop(drag_idx)
        target.insert(drop_idx, item)
        self._refresh_port_trees()
        widget.config(cursor=""); self.drag_data = {"port": None, "iid": None}

    def _refresh_port_trees(self):
        def update_tree(tree, data, entries):
            for i in tree.get_children(): tree.delete(i)
            for t in data:
                row = self.df_master[self.df_master['ticker'] == str(t)]
                name = row.iloc[0]['name'] if not row.empty else "N/A"
                weight = entries.get(t).get() if entries.get(t) else "0.00"
                tree.insert("", "end", values=(t, name, weight))

        update_tree(self.port_a_tree, self.port_a, self.weight_entries_a)
        update_tree(self.port_b_tree, self.port_b, self.weight_entries_b)
        self._refresh_weight_entries("A")
        self._refresh_weight_entries("B")

    def _refresh_weight_entries(self, which):
        tickers = self.port_a if which == "A" else self.port_b
        inner = self.port_a_weights_inner if which == "A" else self.port_b_weights_inner
        entries = self.weight_entries_a if which == "A" else self.weight_entries_b
        
        # Save current values before clear
        old_vals = {t: v.get() for t, v in entries.items()}
        for c in inner.winfo_children(): c.destroy()
        entries.clear()
        
        # Configure columns for inner frame to prevent overlap
        inner.columnconfigure(0, weight=1) # Name label takes space
        inner.columnconfigure(1, weight=0) # Entry stays compact
        
        eq_val = f"{100/(len(tickers) or 1):.2f}"
        for i, t in enumerate(tickers):
            row = self.df_master[self.df_master["ticker"] == str(t)]
            name = row.iloc[0]["name"] if not row.empty else t
            # Display ticker and full name (or longer slice) with padding
            label_text = f"[{t}] {name[:14]}" 
            ttk.Label(inner, text=label_text, style="CardLabel.TLabel").grid(row=i, column=0, sticky="w", padx=(0, 10), pady=2)
            
            var = tk.StringVar(value=old_vals.get(t, eq_val))
            e = ttk.Entry(inner, textvariable=var, width=8, justify="right")
            e.grid(row=i, column=1, sticky="e", pady=2)
            
            entries[t] = var
            var.trace_add("write", lambda *args, w=which: self._update_weight_sum(w))
        self._update_weight_sum(which)

    def _update_weight_sum(self, which):
        ents = self.weight_entries_a if which=="A" else self.weight_entries_b
        try: total = sum(float(v.get() or 0) for v in ents.values())
        except: total = 0
        lbl = self.port_a_sum_label if which=="A" else self.port_b_sum_label
        lbl.config(text=f"합계: {total:.1f}%")

    def _apply_preset(self, which):
        p_name = self.preset_cb.get()
        if p_name == "선택 안함": return
        p_data = backtest.STRATEGY_PRESETS[p_name]
        target = self.port_a if which=="A" else self.port_b
        target.clear()
        for t in p_data.keys(): target.append(t)
        self._refresh_port_trees()
        ents = self.weight_entries_a if which=="A" else self.weight_entries_b
        for t, w in p_data.items():
            if t in ents: ents[t].set(f"{w*100:.1f}")
        self._update_weight_sum(which)

    def _save_config(self):
        config = {
            "start": self.start_var.get(), "initial": self.initial_var.get(), "rf": self.rf_var.get(),
            "rebalance": self.rebalance_var.get(), "benchmark": self.benchmark_display_var.get(),
            "port_a": {t: v.get() for t, v in self.weight_entries_a.items()},
            "port_b": {t: v.get() for t, v in self.weight_entries_b.items()}
        }
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if path:
            with open(path, 'w', encoding='utf-8') as f: json.dump(config, f, indent=4, ensure_ascii=False)
            messagebox.showinfo("저장 완료", "설정이 저장되었습니다.")

    def _load_config(self):
        path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if not path: return
        try:
            with open(path, 'r', encoding='utf-8') as f: cfg = json.load(f)
            self.start_var.set(cfg.get("start", "2018-01-01"))
            self.initial_var.set(cfg.get("initial", "300000000"))
            self.rf_var.set(cfg.get("rf", "0")); self.rebalance_var.set(cfg.get("rebalance", "Monthly"))
            
            # 벤치마크 설정 불러오기 (호환성 유지)
            bench = cfg.get("benchmark", "KOSPI 200")
            if bench in self.benchmark_options:
                self.benchmark_display_var.set(bench)
            else:
                self.benchmark_display_var.set("KOSPI 200")

            self.port_a = list(cfg.get("port_a", {}).keys())
            self.port_b = list(cfg.get("port_b", {}).keys())
            self._refresh_port_trees()
            for t, w in cfg.get("port_a", {}).items():
                if t in self.weight_entries_a: self.weight_entries_a[t].set(w)
            for t, w in cfg.get("port_b", {}).items():
                if t in self.weight_entries_b: self.weight_entries_b[t].set(w)
            self._update_weight_sum("A"); self._update_weight_sum("B")
        except Exception as e: messagebox.showerror("오류", str(e))

    def _normalize_weights(self, which):
        ents = self.weight_entries_a if which=="A" else self.weight_entries_b
        try:
            res = {t: float(v.get() or 0) for t, v in ents.items()}
            s = sum(res.values())
            if s > 0:
                for t, v in ents.items(): v.set(f"{(float(v.get() or 0) / s * 100):.2f}")
        except: pass

    def _refresh_sub_options(self):
        subs = sorted(self.df_master['sub'].dropna().unique().tolist())
        self.sub_cb["values"] = ["전체"] + subs
        if self.sub_var.get() not in self.sub_cb["values"]: self.sub_var.set("전체")

    def _refresh_group_options(self):
        s = self.sub_var.get()
        df = self.df_master
        if s != "전체": df = df[df['sub'] == s]
        groups = sorted(df['group'].dropna().unique().tolist())
        self.group_cb["values"] = ["전체"] + groups
        if self.group_var.get() not in self.group_cb["values"]: self.group_var.set("전체")

    def _refresh_name_list(self):
        df = self.df_master.copy()
        if self.sub_var.get() != "전체": df = df[df['sub'] == self.sub_var.get()]
        if self.group_var.get() != "전체": df = df[df['group'] == self.group_var.get()]
        s = self.search_var.get().strip().lower()
        if s: 
            # 티커 제외, 종목명(name)에서만 검색
            df = df[df['name'].astype(str).str.lower().str.contains(s)]
        
        for i in self.names_tree.get_children(): self.names_tree.delete(i)
        for _, r in df.iterrows():
            self.names_tree.insert("", "end", values=(r.get("ticker", ""), r.get("name", ""), str(r.get("listed", ""))[:10]))

    def _on_sub_changed(self):
        self.group_var.set("전체"); self._refresh_group_options(); self._refresh_name_list()

    def _on_group_changed(self): self._refresh_name_list()

    def _sort_tree(self, col):
        self.sort_reverse = not self.sort_reverse if self.sort_column == col else False
        self.sort_column = col
        items = [(self.names_tree.set(k, col), k) for k in self.names_tree.get_children("")]
        def safe_key(v):
            try: return (0, float(v))
            except: return (1, str(v).lower())
        items.sort(key=lambda x: safe_key(x[0]), reverse=self.sort_reverse)
        for i, (v, k) in enumerate(items): self.names_tree.move(k, '', i)

    def _add_selected_to_port(self, which):
        target = self.port_a if which == "A" else self.port_b
        for iid in self.names_tree.selection():
            t = self.names_tree.item(iid, "values")[0]
            if t not in target: target.append(t)
        self._refresh_port_trees()

    def _remove_selected_from_port(self, which):
        tree = self.port_a_tree if which == "A" else self.port_b_tree
        target = self.port_a if which == "A" else self.port_b
        
        selected_items = tree.selection()
        if not selected_items:
            return

        for item in selected_items:
            ticker = str(tree.item(item)['values'][0])
            if ticker in target:
                target.remove(ticker)
        
        self._refresh_port_trees()

    def _apply_equal_weights(self, which):
        ents = self.weight_entries_a if which=="A" else self.weight_entries_b
        if ents:
            eq = f"{100/len(ents):.2f}"
            for v in ents.values(): v.set(eq)

    def _run(self):
        try:
            if not self.port_a or not self.port_b: raise ValueError("종목을 추가하세요.")
            def get_w(ents):
                res = {t: float(v.get() or 0) for t, v in ents.items()}
                s = sum(res.values())
                if s <= 0: raise ValueError("비중 오류")
                return {t: v/s for t, v in res.items()}
            
            # 벤치마크 표시 이름을 티커로 변환
            bench_ticker = self.benchmark_options.get(self.benchmark_display_var.get(), "069500.KS")
            
            result = backtest.run_backtest(
                backtest.build_df_weights(self.df_master, get_w(self.weight_entries_a), get_w(self.weight_entries_b)),
                start=self.start_var.get(), initial_investment=int(self.initial_var.get()),
                rf=float(self.rf_var.get()), rebalance=self.rebalance_var.get(),
                benchmark_ticker=bench_ticker, show_plots=False, verbose=False
            )
            self._open_results(result)
        except Exception as e: messagebox.showerror("오류", str(e))

    def _open_results(self, result):
        win = tk.Toplevel(self); win.title("Backtest Result"); win.geometry("1200x900"); win.configure(bg=COLOR_BG)
        
        # --- Top Action Bar ---
        top_bar = ttk.Frame(win, padding=(20, 15, 20, 0))
        top_bar.pack(fill="x")
        
        ttk.Label(top_bar, text="Backtest Performance Report", font=FONT_TITLE).pack(side="left")
        
        def export_excel():
            path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if not path: return
            try:
                import io
                img_data = io.BytesIO()
                self.current_fig.savefig(img_data, format='png', dpi=100)
                img_data.seek(0)

                with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
                    result['metrics_compare'].to_excel(writer, sheet_name='Performance_Summary')
                    result['df_weights'].to_excel(writer, sheet_name='Portfolio_Weights', index=False)
                    result['monthly_matrix_a'].to_excel(writer, sheet_name='Monthly_Returns_A')
                    result['monthly_matrix_b'].to_excel(writer, sheet_name='Monthly_Returns_B')
                    
                    # 자산 상관관계 데이터 추가
                    if not result['asset_corr_a'].empty:
                        result['asset_corr_a'].to_excel(writer, sheet_name='Asset_Corr_A')
                    if not result['asset_corr_b'].empty:
                        result['asset_corr_b'].to_excel(writer, sheet_name='Asset_Corr_B')

                    pd.DataFrame({
                        'Port_A': result['asset_values_a'],
                        'Port_B': result['asset_values_b'],
                        'Benchmark': result['asset_values_bench']
                    }).to_excel(writer, sheet_name='Equity_Curve_Data')
                    
                    workbook = writer.book
                    chart_sheet = workbook.add_worksheet('Analysis_Charts')
                    chart_sheet.insert_image('B2', 'backtest_charts.png', {'image_data': img_data})
                
                messagebox.showinfo("완료", "자산 상관관계를 포함한 전체 보고서가 저장되었습니다.")
            except Exception as e:
                messagebox.showerror("오류", str(e))

        ttk.Button(top_bar, text="📥 전체 결과 보고서(엑셀) 내보내기", style="Success.TButton", command=export_excel).pack(side="right")

        nb = ttk.Notebook(win); nb.pack(fill="both", expand=True, padx=20, pady=(10, 20))

        # Tab 1: Stats
        t1 = ttk.Frame(nb, padding=15); nb.add(t1, text="Performance Stats")
        tree = ttk.Treeview(t1, columns=("Metric", "Port A", "Port B", "Benchmark"), show="headings")
        for c in tree["columns"]: tree.heading(c, text=c); tree.column(c, width=150, anchor="center")
        tree.pack(fill="both", expand=True)
        for idx, row in result["metrics_compare"].iterrows():
            tree.insert("", "end", values=(idx, row.get("Port A", "-"), row.get("Port B", "-"), row.get("Benchmark", "-")))

        # Tab 2: Monthly Returns Heatmap (Combined A & B)
        t_monthly = ttk.Frame(nb, padding=15); nb.add(t_monthly, text="Monthly Returns")
        fig_m = Figure(figsize=(12, 6))
        axs_m = fig_m.subplots(1, 2)
        
        def draw_monthly_heatmap(ax, df, title):
            if df.empty:
                ax.text(0.5, 0.5, "No Data", ha='center', va='center')
                return
            im = ax.imshow(df.values, cmap='RdYlGn', vmin=-5, vmax=5, aspect='auto')
            ax.set_title(title, fontweight='bold')
            ax.set_xticks(np.arange(len(df.columns))); ax.set_yticks(np.arange(len(df.index)))
            ax.set_xticklabels([f"{m}M" for m in df.columns], fontsize=8)
            ax.set_yticklabels(df.index, fontsize=8)
            # 수치 표시
            for i in range(len(df.index)):
                for j in range(len(df.columns)):
                    val = df.iloc[i, j]
                    if pd.notna(val):
                        ax.text(j, i, f"{val:.1f}", ha="center", va="center", color="black", fontsize=7)
            fig_m.colorbar(im, ax=ax, fraction=0.046, pad=0.04)

        draw_monthly_heatmap(axs_m[0], result['monthly_matrix_a'], "Port A: Monthly Returns (%)")
        draw_monthly_heatmap(axs_m[1], result['monthly_matrix_b'], "Port B: Monthly Returns (%)")
        fig_m.tight_layout()
        canvas_m = FigureCanvasTkAgg(fig_m, master=t_monthly); canvas_m.draw()
        canvas_m.get_tk_widget().pack(fill="both", expand=True)

        # Tab 3: Asset Correlation Analysis
        t_corr = ttk.Frame(nb, padding=15); nb.add(t_corr, text="Asset Correlation")
        fig_corr = Figure(figsize=(10, 5))
        axs_corr = fig_corr.subplots(1, 2)
        
        def draw_heatmap(ax, df, title):
            if df.empty:
                ax.text(0.5, 0.5, "No Data", ha='center', va='center')
                return
            im = ax.imshow(df.values, cmap='RdYlGn', vmin=-1, vmax=1)
            ax.set_title(title, fontweight='bold')
            ax.set_xticks(np.arange(len(df.columns))); ax.set_yticks(np.arange(len(df.index)))
            # 레이블을 종목명으로 표시 (이미 backtest.py에서 처리됨)
            ax.set_xticklabels(df.columns, rotation=45, ha='right', fontsize=7)
            ax.set_yticklabels(df.index, fontsize=7)
            # 상관계수 숫자 표시
            for i in range(len(df.index)):
                for j in range(len(df.columns)):
                    ax.text(j, i, f"{df.iloc[i, j]:.2f}", ha="center", va="center", color="black", fontsize=7)
            fig_corr.colorbar(im, ax=ax, fraction=0.046, pad=0.04)

        draw_heatmap(axs_corr[0], result['asset_corr_a'], "Port A: Asset Correlation")
        draw_heatmap(axs_corr[1], result['asset_corr_b'], "Port B: Asset Correlation")
        fig_corr.tight_layout()
        canvas_corr = FigureCanvasTkAgg(fig_corr, master=t_corr); canvas_corr.draw()
        canvas_corr.get_tk_widget().pack(fill="both", expand=True)

        # Tab 5: Charts
        t3 = ttk.Frame(nb, padding=10); nb.add(t3, text="Analysis Charts")
        fig = Figure(figsize=(10, 10)); axs = fig.subplots(3, 1)
        self.current_fig = fig  # 엑셀 저장을 위해 참조 저장
        
        # Cumulative Equity Curve (기존 코드와 동일)
        axs[0].plot(result["asset_values_a"], label="Port A", color=COLOR_ACCENT, lw=2)
        axs[0].plot(result["asset_values_b"], label="Port B", color="#E67E22", lw=2)
        axs[0].plot(result["asset_values_bench"], label="Benchmark", ls="--", color="gray")
        axs[0].set_title("Cumulative Equity Curve", fontsize=11, fontweight='bold')
        axs[0].legend(); axs[0].grid(True, ls=":")
        axs[0].set_yscale('linear')
        from matplotlib.ticker import FuncFormatter
        axs[0].yaxis.set_major_formatter(FuncFormatter(lambda x, p: format(int(x), ',')))
        
        # 12M Rolling Returns
        axs[1].plot(result['rolling_12m_a'], label="Port A 1Y Rolling", color=COLOR_ACCENT)
        axs[1].plot(result['rolling_12m_b'], label="Port B 1Y Rolling", color="#E67E22")
        axs[1].axhline(0, color="black", lw=1)
        axs[1].set_title("12M Rolling Returns (%)", fontsize=11, fontweight='bold')
        axs[1].legend(); axs[1].grid(True, ls=":")
        axs[1].set_yscale('linear')
        
        # Drawdown
        axs[2].fill_between(result["drawdown_a"].index, result["drawdown_a"], color=COLOR_ACCENT, alpha=0.2)
        axs[2].plot(result["drawdown_a"], color=COLOR_ACCENT, label="Port A DD")
        axs[2].plot(result["drawdown_b"], color="#E67E22", label="Port B DD")
        axs[2].set_title("Drawdown (%)", fontsize=11, fontweight='bold')
        axs[2].grid(True, ls=":"); axs[2].legend()
        axs[2].set_yscale('linear')
        
        fig.tight_layout(); canvas = FigureCanvasTkAgg(fig, master=t3); canvas.draw(); canvas.get_tk_widget().pack(fill="both", expand=True)

if __name__ == "__main__": BacktestGUI().mainloop()
