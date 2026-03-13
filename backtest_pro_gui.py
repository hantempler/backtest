import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
import json
import io
import threading
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import backtest_proxy
import matplotlib.pyplot as plt
from matplotlib import patheffects

# --- Matplotlib 한글 폰트 설정 ---
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

# --- UI Constants ---
COLOR_BG = "#F8F9FA"
COLOR_CARD = "#FFFFFF"
COLOR_PRIMARY = "#2C3E50"
COLOR_ACCENT = "#3498DB"
COLOR_SUCCESS = "#27AE60"
COLOR_BORDER = "#DEE2E6"
FONT_MAIN = ("Segoe UI", 10)
FONT_BOLD = ("Segoe UI", 10, "bold")
FONT_TITLE = ("Segoe UI", 12, "bold")

class ProBacktestGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("K-Global 하이브리드 자산배분 시뮬레이터 v1.5 (적립식 지원)")
        self.geometry("1400x950")
        self.configure(bg=COLOR_BG)

        self.port_a = {}
        self.port_b = {}
        self.start_var = tk.StringVar(value="2000-01-01")
        self.initial_var = tk.StringVar(value="100,000,000")
        self.monthly_contrib_var = tk.StringVar(value="0")
        self.bench_var = tk.StringVar(value="S&P500 (SPY)")
        self.rebalance_var = tk.StringVar(value="매월")
        self.currency_var = tk.StringVar(value="KRW")
        self.fee_var = tk.StringVar(value="0.1")
        self.threshold_var = tk.StringVar(value="5.0")
        self.custom_ticker_var = tk.StringVar()
        
        # --- Excel Data & Filtering Variables ---
        self.df_master = pd.read_excel("etf_data.xlsx")
        
        # Precise column mapping
        mapping = {
            '단축코드': 'ticker',
            '한글종목약명': 'name',
            '기초시장분류': 'sub',
            '기초자산분류': 'group',
            '상장일': 'listed'
        }
        self.df_master = self.df_master.rename(columns=mapping)
        self.df_master['sub'] = '국내상장ETF'
        
        # Merge Hardcoded Proxy Assets
        proxy_rows = []
        for cat, assets in backtest_proxy.ASSET_UNIVERSE.items():
            for t, n in assets.items():
                proxy_rows.append({
                    'ticker': t, 'name': n.split(' (')[0], 
                    'sub': 'Global Strategy', 'group': cat, 
                    'listed': 'Fixed'
                })
        df_proxy = pd.DataFrame(proxy_rows)
        self.df_master = pd.concat([df_proxy, self.df_master], ignore_index=True)
        
        if 'sub' not in self.df_master.columns: self.df_master['sub'] = '미분류'
        if 'group' not in self.df_master.columns: self.df_master['group'] = '미분류'
        if 'name' not in self.df_master.columns: self.df_master['name'] = '이름없음'
        if 'ticker' not in self.df_master.columns: 
             possible_ticker_cols = [c for c in self.df_master.columns if '코드' in str(c)]
             if possible_ticker_cols: self.df_master['ticker'] = self.df_master[possible_ticker_cols[0]]
             else: self.df_master['ticker'] = self.df_master.index.astype(str)
        
        self.df_master['ticker'] = self.df_master['ticker'].apply(lambda x: str(x).zfill(6) if str(x).isdigit() else str(x))
        
        # Update Global Ticker Name Map
        for _, row in self.df_master.iterrows():
            t = str(row['ticker'])
            if not t.endswith(".KS") and not any(t in assets for assets in backtest_proxy.ASSET_UNIVERSE.values()): 
                t = f"{t}.KS"
            backtest_proxy.GLOBAL_TICKER_NAME_MAP[t] = str(row['name'])
        
        self.sub_var = tk.StringVar(value="전체")
        self.group_var = tk.StringVar(value="전체")
        self.search_var = tk.StringVar(value="")
        self.sort_col = "ticker"; self.sort_asc = True
        self.drag_data = {"source": None, "ticker": None, "port": None, "iid": None}

        self._apply_styles()
        self._build_ui()
        self._refresh_sub_options()
        self._refresh_asset_tree()

    def _apply_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure(".", font=FONT_MAIN, background=COLOR_BG, foreground=COLOR_PRIMARY)
        style.configure("Card.TFrame", background=COLOR_CARD, relief="flat")
        style.configure("Header.TLabel", font=FONT_TITLE, background=COLOR_CARD)
        style.configure("CardLabel.TLabel", background=COLOR_CARD, font=FONT_BOLD)
        style.configure("Accent.TButton", font=FONT_BOLD, foreground="white", background=COLOR_ACCENT)
        style.configure("Success.TButton", font=FONT_BOLD, foreground="white", background=COLOR_SUCCESS)
        style.configure("Treeview", font=FONT_MAIN, rowheight=28)
        style.configure("Treeview.Heading", font=FONT_BOLD)

    def _build_ui(self):
        root = ttk.Frame(self, padding=20); root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=3); root.columnconfigure(1, weight=7) 
        sidebar = ttk.Frame(root); sidebar.grid(row=0, column=0, sticky="nsew", padx=(0, 15))

        # 1. Presets
        cp = ttk.Frame(sidebar, style="Card.TFrame", padding=15); cp.pack(fill="x", pady=(0, 15))
        ttk.Label(cp, text="전략 프리셋 (Presets)", style="Header.TLabel").pack(anchor="w", pady=(0, 10))
        self.preset_cb = ttk.Combobox(cp, values=["선택 안함"] + list(backtest_proxy.STRATEGY_PRESETS.keys()), state="readonly")
        self.preset_cb.set("선택 안함"); self.preset_cb.pack(fill="x", pady=(0, 10))
        btn_pf = ttk.Frame(cp, style="Card.TFrame"); btn_pf.pack(fill="x")
        ttk.Button(btn_pf, text="A에 적용", command=lambda: self._apply_preset("A")).pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(btn_pf, text="B에 적용", command=lambda: self._apply_preset("B")).pack(side="left", fill="x", expand=True)

        # 2. Asset Universe
        c1 = ttk.Frame(sidebar, style="Card.TFrame", padding=15); c1.pack(fill="both", expand=True, pady=(0, 15))
        ttk.Label(c1, text="자산 유니버스 (드래그하여 추가)", style="Header.TLabel").pack(anchor="w", pady=(0, 10))
        filter_f = ttk.Frame(c1, style="Card.TFrame"); filter_f.pack(fill="x", pady=(0, 10))
        ttk.Label(filter_f, text="대분류:", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")
        self.sub_cb = ttk.Combobox(filter_f, textvariable=self.sub_var, state="readonly", width=15)
        self.sub_cb.grid(row=0, column=1, sticky="ew", padx=(5, 0), pady=2); self.sub_cb.bind("<<ComboboxSelected>>", self._on_sub_changed)
        ttk.Label(filter_f, text="중분류:", style="CardLabel.TLabel").grid(row=1, column=0, sticky="w")
        self.group_cb = ttk.Combobox(filter_f, textvariable=self.group_var, state="readonly", width=15)
        self.group_cb.grid(row=1, column=1, sticky="ew", padx=(5, 0), pady=2); self.group_cb.bind("<<ComboboxSelected>>", lambda e: self._refresh_asset_tree())
        ttk.Label(filter_f, text="검색어:", style="CardLabel.TLabel").grid(row=2, column=0, sticky="w")
        self.search_ent = ttk.Entry(filter_f, textvariable=self.search_var, width=17)
        self.search_ent.grid(row=2, column=1, sticky="ew", padx=(5, 0), pady=2); self.search_var.trace_add("write", lambda *args: self._refresh_asset_tree())
        filter_f.columnconfigure(1, weight=1)

        self.asset_tree = ttk.Treeview(c1, columns=("ticker", "name", "listed"), show="headings", height=12)
        for col, head in [("ticker", "티커"), ("name", "자산명"), ("listed", "상장일")]:
            self.asset_tree.heading(col, text=head, command=lambda c=col: self._on_header_click(c))
        self.asset_tree.column("ticker", width=40, anchor="center"); self.asset_tree.column("name", width=160); self.asset_tree.column("listed", width=40, anchor="center")
        self.asset_tree.pack(fill="both", expand=True)
        self.asset_tree.bind("<Button-1>", self._start_universe_drag); self.asset_tree.bind("<B1-Motion>", self._drag_motion); self.asset_tree.bind("<ButtonRelease-1>", self._drop)
        btn_f = ttk.Frame(c1, style="Card.TFrame"); btn_f.pack(fill="x", pady=(10, 0))
        ttk.Button(btn_f, text="+ A추가", width=7, command=lambda: self._add_to_port("A")).pack(side="left", padx=(0, 5))
        ttk.Button(btn_f, text="+ B추가", width=7, command=lambda: self._add_to_port("B")).pack(side="left")
        ttk.Label(btn_f, text="직접입력:", style="CardLabel.TLabel").pack(side="left", padx=(10, 5))
        ttk.Entry(btn_f, textvariable=self.custom_ticker_var, width=8).pack(side="left")
        ttk.Button(btn_f, text="등록", width=5, command=self._add_custom).pack(side="left", padx=5)

        main_area = ttk.Frame(root); main_area.grid(row=0, column=1, sticky="nsew"); main_area.rowconfigure(2, weight=1)
        
        # 3. Configuration (Moved to main area top)
        c2 = ttk.Frame(main_area, style="Card.TFrame", padding=15); c2.pack(fill="x", pady=(0, 15))
        ttk.Label(c2, text="시뮬레이션 설정 (Simulation Settings)", style="Header.TLabel").pack(anchor="w", pady=(0, 10))
        grid_f = ttk.Frame(c2, style="Card.TFrame"); grid_f.pack(fill="x")
        
        self.bench_map = {"S&P500 (SPY)": "SPY", "나스닥100 (QQQ)": "QQQ", "한국주식 (EWY)": "EWY", "미국주식-전체 (VTI)": "VTI", "코스피 200 (KODEX)": "069500.KS"}
        self.bench_var.set("S&P500 (SPY)") 
        
        labels = [
            ("시작일", self.start_var), 
            ("초기 투자금", self.initial_var), 
            ("월 적립금", self.monthly_contrib_var), 
            ("벤치마크", self.bench_var), 
            ("리밸런싱", self.rebalance_var), 
            ("기준 통화", self.currency_var),
            ("거래 수수료(%)", self.fee_var),
            ("리밸런싱 임계치(%)", self.threshold_var)
        ]
        
        # 가로 배치를 위한 그리드 구성 (4열: 라벨, 위젯, 라벨, 위젯)
        for i, (txt, var) in enumerate(labels):
            row = i // 4
            col = (i % 4) * 2
            ttk.Label(grid_f, text=txt, style="CardLabel.TLabel").grid(row=row, column=col, sticky="w", padx=(15, 5), pady=5)
            
            if txt == "벤치마크":
                cb = ttk.Combobox(grid_f, textvariable=var, values=list(self.bench_map.keys()), state="readonly", width=18)
                cb.grid(row=row, column=col+1, sticky="ew", pady=5)
            elif txt == "리밸런싱":
                cb = ttk.Combobox(grid_f, textvariable=var, values=["매월", "분기별", "매년", "비중이탈시", "안함"], state="readonly", width=18)
                cb.grid(row=row, column=col+1, sticky="ew", pady=5)
            elif txt == "기준 통화":
                cb = ttk.Combobox(grid_f, textvariable=var, values=["KRW", "USD"], state="readonly", width=18)
                cb.grid(row=row, column=col+1, sticky="ew", pady=5)
            else:
                ent = ttk.Entry(grid_f, textvariable=var, width=15)
                ent.grid(row=row, column=col+1, sticky="ew", pady=5)
        
        grid_f.columnconfigure((1, 3, 5, 7), weight=1)

        manage_f = ttk.Frame(main_area, style="Card.TFrame", padding=10); manage_f.pack(fill="x", pady=(0, 15))
        ttk.Label(manage_f, text="포트폴리오 구성 및 관리", style="Header.TLabel").pack(side="left", padx=5)
        ttk.Button(manage_f, text="📂 불러오기", width=12, command=self._load_config).pack(side="right", padx=5)
        ttk.Button(manage_f, text="💾 현재 설정 저장", width=15, style="Accent.TButton", command=self._save_config).pack(side="right", padx=5)

        ports_f = ttk.Frame(main_area); ports_f.pack(fill="both", expand=True)
        ports_f.columnconfigure(0, weight=1); ports_f.columnconfigure(1, weight=1)

        def build_port_box(parent, title, which):
            frame = ttk.Frame(parent, style="Card.TFrame", padding=15)
            ttk.Label(frame, text=title, style="Header.TLabel").pack(anchor="w", pady=(0, 10))
            tree = ttk.Treeview(frame, columns=("ticker", "name", "weight"), show="headings", height=10)
            tree.heading("ticker", text="티커"); tree.heading("name", text="자산명"); tree.heading("weight", text="비중(%)")
            tree.column("ticker", width=70, anchor="center", stretch=False); tree.column("name", width=200, anchor="w", stretch=True); tree.column("weight", width=60, anchor="center", stretch=False)
            tree.pack(fill="x")
            tree.bind("<Button-1>", self._start_port_drag); tree.bind("<B1-Motion>", self._drag_motion); tree.bind("<ButtonRelease-1>", self._drop)
            scroll_f = ttk.Frame(frame, style="Card.TFrame"); scroll_f.pack(fill="both", expand=True, pady=(10, 0))
            canvas = tk.Canvas(scroll_f, bg="white", highlightthickness=0, height=250)
            scrollbar = ttk.Scrollbar(scroll_f, orient="vertical", command=canvas.yview)
            edit_inner = ttk.Frame(canvas, style="Card.TFrame")
            edit_inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
            canvas.create_window((0, 0), window=edit_inner, anchor="nw"); canvas.configure(yscrollcommand=scrollbar.set)
            canvas.pack(side="left", fill="both", expand=True); scrollbar.pack(side="right", fill="y")
            for w in [frame, tree, scroll_f, canvas, edit_inner]: w.bind("<ButtonRelease-1>", self._drop, add="+")
            btn_box = ttk.Frame(frame, style="Card.TFrame"); btn_box.pack(fill="x", pady=(10, 0))
            ttk.Button(btn_box, text="삭제", width=5, command=lambda: self._delete_asset(which)).pack(side="left")
            ttk.Button(btn_box, text="동일비중", width=8, command=lambda: self._equal_weight(which)).pack(side="left", padx=5)
            ttk.Button(btn_box, text="초기화", width=7, command=lambda: self._reset_weights(which)).pack(side="left")
            sum_lbl = ttk.Label(btn_box, text="합계: 0%", style="CardLabel.TLabel"); sum_lbl.pack(side="right")
            if which == "A": self.tree_a, self.edit_a, self.sum_a, self.box_a = tree, edit_inner, sum_lbl, frame
            else: self.tree_b, self.edit_b, self.sum_b, self.box_b = tree, edit_inner, sum_lbl, frame
            return frame

        build_port_box(ports_f, "포트폴리오 A (전략)", "A").grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        build_port_box(ports_f, "포트폴리오 B (비교)", "B").grid(row=0, column=1, sticky="nsew")
        ttk.Button(main_area, text="🚀 장기 백테스트 실행 (BACKTEST)", style="Success.TButton", command=self._run).pack(fill="x", pady=(15, 0), ipady=12)

    def _start_universe_drag(self, event):
        iid = self.asset_tree.identify_row(event.y)
        if iid: vals = self.asset_tree.item(iid)['values']; self.drag_data = {"source": "universe", "ticker": str(vals[0]), "port": None, "iid": iid}; self.asset_tree.config(cursor="hand2")
    def _start_port_drag(self, event):
        widget = event.widget; iid = widget.identify_row(event.y)
        if iid: self.drag_data = {"source": "portfolio", "ticker": str(widget.item(iid)['values'][0]), "port": "A" if widget == self.tree_a else "B", "iid": iid}; widget.config(cursor="hand2")
    def _drag_motion(self, event): pass
    def _drop(self, event):
        if not self.drag_data["ticker"]: return
        x, y = event.x_root, event.y_root; target_widget = self.winfo_containing(x, y)
        def is_descendant(child, parent):
            while child:
                if child == parent: return True
                child = child.master
            return False
        drop_target = "A" if is_descendant(target_widget, self.box_a) else ("B" if is_descendant(target_widget, self.box_b) else None)
        source, ticker = self.drag_data["source"], self.drag_data["ticker"]
        if source == "universe" and drop_target:
            is_global = False
            for assets in backtest_proxy.ASSET_UNIVERSE.values():
                if ticker in assets: is_global = True; break
            if not is_global and not ticker.endswith(".KS"): ticker = f"{ticker}.KS"
            target_dict = self.port_a if drop_target == "A" else self.port_b
            if ticker not in target_dict: target_dict[ticker] = 0.0; self._refresh_port_ui(drop_target)
        elif source == "portfolio" and drop_target and drop_target != self.drag_data["port"]:
            target_dict = self.port_a if drop_target == "A" else self.port_b
            if ticker not in target_dict:
                source_dict = self.port_a if self.drag_data["port"] == "A" else self.port_b
                target_dict[ticker] = source_dict.get(ticker, 0.0); self._refresh_port_ui(drop_target)
        elif source == "portfolio" and drop_target == self.drag_data["port"]:
            tree = self.tree_a if drop_target == "A" else self.tree_b; rel_y = y - tree.winfo_rooty(); drop_iid = tree.identify_row(rel_y)
            if drop_iid and self.drag_data["iid"] != drop_iid:
                target = self.port_a if drop_target == "A" else self.port_b; items = list(target.items())
                idx1 = [i[0] for i in items].index(ticker); idx2 = [i[0] for i in items].index(str(tree.item(drop_iid)['values'][0]))
                val = items.pop(idx1); items.insert(idx2, val)
                if drop_target == "A": self.port_a = dict(items)
                else: self.port_b = dict(items)
                self._refresh_port_ui(drop_target)
        self.asset_tree.config(cursor=""); self.tree_a.config(cursor=""); self.tree_b.config(cursor=""); self.drag_data = {"source": None, "ticker": None, "port": None, "iid": None}

    def _apply_preset(self, which):
        p_name = self.preset_cb.get()
        if p_name != "선택 안함":
            target = self.port_a if which == "A" else self.port_b; target.clear()
            for t, w in backtest_proxy.STRATEGY_PRESETS[p_name].items(): target[str(t)] = w
            self._refresh_port_ui(which)
    def _add_to_port(self, which):
        sel = self.asset_tree.selection()
        if sel:
            item = self.asset_tree.item(sel[0]); ticker = str(item['values'][0])
            is_global = False
            for assets in backtest_proxy.ASSET_UNIVERSE.values():
                if ticker in assets: is_global = True; break
            if not is_global and not ticker.endswith(".KS"): ticker = f"{ticker}.KS"
            target = self.port_a if which == "A" else self.port_b
            if ticker not in target: target[ticker] = 0.0; self._refresh_port_ui(which)
    def _add_custom(self):
        t = self.custom_ticker_var.get().strip().upper()
        if t:
            if t not in self.port_a: self.port_a[t] = 0.0; self._refresh_port_ui("A")
            self.custom_ticker_var.set("")
    def _refresh_port_ui(self, which):
        tree, edit, data, sum_lbl = (self.tree_a, self.edit_a, self.port_a, self.sum_a) if which == "A" else (self.tree_b, self.edit_b, self.port_b, self.sum_b)
        for i in tree.get_children(): tree.delete(i)
        for c in edit.winfo_children(): c.destroy()
        edit.columnconfigure(0, minsize=70); edit.columnconfigure(1, minsize=200); edit.columnconfigure(2, minsize=60)
        ttk.Label(edit, text="티커", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w", padx=5)
        ttk.Label(edit, text="자산명", style="CardLabel.TLabel").grid(row=0, column=1, sticky="w", padx=5)
        ttk.Label(edit, text="비중(%)", style="CardLabel.TLabel").grid(row=0, column=2, sticky="e", padx=5)
        total_w = 0
        for i, (t, w) in enumerate(data.items()):
            name = backtest_proxy.get_asset_name(t).split(' (')[0]
            tree.insert("", "end", values=(t, name, f"{w*100:.1f}")); total_w += w
            ttk.Label(edit, text=t, style="CardLabel.TLabel").grid(row=i+1, column=0, sticky="w", padx=5)
            ttk.Label(edit, text=name[:18], style="CardLabel.TLabel").grid(row=i+1, column=1, sticky="w", padx=5)
            var = tk.StringVar(value=f"{w*100:.1f}"); ent = ttk.Entry(edit, textvariable=var, width=8, justify="right")
            ent.grid(row=i+1, column=2, sticky="e", padx=5, pady=2); var.trace_add("write", lambda *args, t=t, v=var, w=which: self._update_weight(w, t, v))
        sum_lbl.config(text=f"합계: {total_w*100:.1f}%")
    def _update_weight(self, which, ticker, var):
        try:
            val = float(var.get() or 0) / 100
            if which == "A": self.port_a[ticker] = val
            else: self.port_b[ticker] = val
            data = self.port_a if which == "A" else self.port_b; sum_lbl = self.sum_a if which == "A" else self.sum_b
            sum_lbl.config(text=f"합계: {sum(data.values())*100:.1f}%")
        except: pass
    def _save_config(self):
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON 파일", "*.json")])
        if path:
            try:
                config = {
                    "start": self.start_var.get(), 
                    "initial": self.initial_var.get(), 
                    "contrib": self.monthly_contrib_var.get(), 
                    "bench": self.bench_var.get(), 
                    "reb": self.rebalance_var.get(), 
                    "cur": self.currency_var.get(), 
                    "fee": self.fee_var.get(),
                    "threshold": self.threshold_var.get(),
                    "a": self.port_a, 
                    "b": self.port_b
                }
                with open(path, 'w', encoding='utf-8') as f: json.dump(config, f, indent=4, ensure_ascii=False)
                messagebox.showinfo("성공", "포트폴리오 설정이 저장되었습니다.")
            except Exception as e: messagebox.showerror("에러", str(e))
    def _load_config(self):
        path = filedialog.askopenfilename(filetypes=[("JSON 파일", "*.json")])
        if path:
            try:
                with open(path, 'r', encoding='utf-8') as f: cfg = json.load(f)
                self.start_var.set(cfg.get("start", "2005-01-01")); self.initial_var.set(cfg.get("initial", "100,000,000")); self.monthly_contrib_var.set(cfg.get("contrib", "0"))
                self.bench_var.set(cfg.get("bench", "S&P500 (SPY)")); self.rebalance_var.set(cfg.get("reb", "매월")); self.currency_var.set(cfg.get("cur", "KRW"))
                self.fee_var.set(cfg.get("fee", "0.1")); self.threshold_var.set(cfg.get("threshold", "5.0"))
                self.port_a, self.port_b = cfg.get("a", {}), cfg.get("b", {}); self._refresh_port_ui("A"); self._refresh_port_ui("B")
            except Exception as e: messagebox.showerror("에러", str(e))
    def _delete_asset(self, which):
        tree, data = (self.tree_a, self.port_a) if which == "A" else (self.tree_b, self.port_b)
        for item in tree.selection():
            ticker = str(tree.item(item)['values'][0])
            if ticker in data: del data[ticker]
        self._refresh_port_ui(which)
    def _equal_weight(self, which):
        data = self.port_a if which == "A" else self.port_b
        if data:
            eq = 1.0 / len(data); [data.__setitem__(t, eq) for t in data]; self._refresh_port_ui(which)
    def _reset_weights(self, which):
        data = self.port_a if which == "A" else self.port_b
        if data: [data.__setitem__(t, 0.0) for t in data]; self._refresh_port_ui(which)

    def _run(self):
        try:
            has_a = bool(self.port_a); has_b = bool(self.port_b)
            if not has_a and not has_b: raise ValueError("최소 하나의 포트폴리오를 구성해야 합니다.")
            if has_a and abs(sum(self.port_a.values()) - 1.0) > 0.05: raise ValueError("포트폴리오 A 비중 합계는 100%여야 합니다.")
            if has_b and abs(sum(self.port_b.values()) - 1.0) > 0.05: raise ValueError("포트폴리오 B 비중 합계는 100%여야 합니다.")
            bench_ticker = self.bench_map.get(self.bench_var.get(), "SPY")
            reb_map = {"매월": "Monthly", "분기별": "Quarterly", "매년": "Yearly", "비중이탈시": "Threshold", "안함": "None"}
            rebalance_eng = reb_map.get(self.rebalance_var.get(), "Monthly")
            
            fee_rate = float(self.fee_var.get() or 0) / 100
            threshold_rate = float(self.threshold_var.get() or 0) / 100
            
            res = backtest_proxy.run_pro_backtest(
                self.port_a, self.port_b, 
                start=self.start_var.get(), 
                initial_investment=int(self.initial_var.get().replace(',', '') or 100000000), 
                benchmark_ticker=bench_ticker, 
                rebalance=rebalance_eng, 
                base_currency=self.currency_var.get(), 
                monthly_contribution=float(self.monthly_contrib_var.get().replace(',', '') or 0),
                transaction_fee=fee_rate,
                rebalance_threshold=threshold_rate
            )
            self._show_results(res)
        except Exception as e: messagebox.showerror("시뮬레이션 에러", str(e))

    def _show_results(self, res):
        win = tk.Toplevel(self); win.title("백테스트 상세 리포트"); win.geometry("1300x950"); win.configure(bg=COLOR_BG)
        top = ttk.Frame(win, padding=(20, 10)); top.pack(fill="x")
        curr = self.currency_var.get(); ttk.Label(top, text=f"투자 성과 분석 ({curr} 기준)", font=FONT_TITLE).pack(side="left")
        def export_excel():
            path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 파일", "*.xlsx")])
            if path:
                try:
                    with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
                        res['metrics'].to_excel(writer, sheet_name='성과요약')
                        for p_idx, p_data in [('A', self.port_a), ('B', self.port_b)]:
                            if not p_data: continue
                            pd.DataFrame([{"티커": t, "자산명": backtest_proxy.get_asset_name(t).split(' (')[0], "비중 (%)": w * 100} for t, w in p_data.items()]).to_excel(writer, sheet_name=f'구성비중_{p_idx}', index=False)
                        for p_idx, con_data in [('A', res['contrib_a']), ('B', res['contrib_b'])]:
                            if con_data is None or con_data.empty or con_data.sum() == 0: continue
                            pd.DataFrame([{"티커": t, "자산명": backtest_proxy.get_asset_name(t).split(' (')[0], "기여도 (%)": val * 100} for t, val in con_data.items() if val != 0]).sort_values(by="기여도 (%)", ascending=False).to_excel(writer, sheet_name=f'수익기여도_{p_idx}', index=False)
                        for p_idx, p_data in [('A', self.port_a), ('B', self.port_b)]:
                            if not p_data: continue
                            valid_tickers = [t for t in p_data.keys() if t in res['raw_returns'].columns]
                            if not valid_tickers: continue
                            detail_dfs = []
                            fx_data = res['raw_prices']['KRW=X'] if 'KRW=X' in res['raw_prices'].columns else pd.Series(1.0, index=res['raw_prices'].index)
                            for t in valid_tickers:
                                t_df = pd.DataFrame(index=res['raw_returns'].index); t_df[f'{t}_시장가격'] = res['raw_prices'][t] if t in res['raw_prices'].columns else np.nan; t_df[f'환율'] = fx_data; t_df[f'{t}_지수화가치'] = (1 + res['raw_returns'][t]).cumprod() * 100_000_000; detail_dfs.append(t_df)
                            pd.concat(detail_dfs, axis=1).to_excel(writer, sheet_name=f'상세데이터_{p_idx}')
                        if not res['monthly_a'].empty: res['monthly_a'].to_excel(writer, sheet_name='월간수익률_A')
                        if not res['monthly_b'].empty: res['monthly_b'].to_excel(writer, sheet_name='월간수익률_B')
                        if not res['corr_a'].empty: res['corr_a'].to_excel(writer, sheet_name='상관계수_A')
                        if not res['corr_b'].empty: res['corr_b'].to_excel(writer, sheet_name='상관계수_B')
                    messagebox.showinfo("성공", "엑셀 리포트 저장이 완료되었습니다.")
                except Exception as e: messagebox.showerror("에러", str(e))
        ttk.Button(top, text="📥 엑셀 다운로드 (Export)", style="Success.TButton", command=export_excel).pack(side="right")
        nb = ttk.Notebook(win); nb.pack(fill="both", expand=True, padx=20, pady=10)
        def create_scroll_tab(notebook, title):
            frame = ttk.Frame(notebook); notebook.add(frame, text=title); canvas = tk.Canvas(frame, bg=COLOR_BG, highlightthickness=0); scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview); scrollable_inner = ttk.Frame(canvas, style="Card.TFrame")
            scrollable_inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))); canvas_window = canvas.create_window((0, 0), window=scrollable_inner, anchor="nw")
            canvas.bind("<Configure>", lambda e: (canvas.itemconfig(canvas_window, width=e.width), canvas.coords(canvas_window, e.width/2, 0)), add="+"); canvas.configure(yscrollcommand=scrollbar.set); canvas.pack(side="left", fill="both", expand=True); scrollbar.pack(side="right", fill="y")
            def _on_mw(e): canvas.yview_scroll(int(-1*(e.delta/120)), "units")
            canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mw)); canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
            return scrollable_inner
        UNIFIED_FIG = (15, 15); t1_inner = create_scroll_tab(nb, "1. 성과 통계"); comp_frame = ttk.Frame(t1_inner, style="Card.TFrame"); comp_frame.pack(fill="x", padx=30, pady=(20, 0)); comp_frame.columnconfigure(0, weight=1); comp_frame.columnconfigure(1, weight=1)
        def build_comp_mini_table(parent, title, data, col):
            f = ttk.LabelFrame(parent, text=title, padding=10); f.grid(row=0, column=col, sticky="nsew", padx=10); t = ttk.Treeview(f, columns=("name", "weight"), show="headings", height=6); t.heading("name", text="자산명"); t.heading("weight", text="비중 (%)"); t.column("name", width=180); t.column("weight", width=80, anchor="center"); t.pack(fill="both", expand=True)
            if data: [t.insert("", "end", values=(backtest_proxy.get_asset_name(tk).split(' (')[0], f"{w*100:.1f}%")) for tk, w in data.items()]
            else: t.insert("", "end", values=("구성 없음", "-"))
            return t
        build_comp_mini_table(comp_frame, "포트폴리오 A 구성", self.port_a, 0); build_comp_mini_table(comp_frame, "포트폴리오 B 구성", self.port_b, 1)
        ttk.Label(t1_inner, text="핵심 성과 지표 (Key Metrics)", font=FONT_TITLE, background=COLOR_CARD).pack(pady=(30, 0))
        tree = ttk.Treeview(t1_inner, columns=("Metric", "Port A", "Port B", "Benchmark"), show="headings", height=15)
        for c in tree["columns"]: tree.heading(c, text=c); tree.column(c, anchor="center", width=250)
        tree.pack(pady=(10, 30), padx=30, anchor="n"); [tree.insert("", "end", values=(idx, row['Port A'], row['Port B'], row['Benchmark'])) for idx, row in res['metrics'].iterrows()]
        t2_inner = create_scroll_tab(nb, "2. 수익률 및 리스크"); fig = Figure(figsize=UNIFIED_FIG); axs = fig.subplots(2, 1); bench = self.bench_var.get()
        if res['asset_values_a'] is not None: axs[0].plot(res['asset_values_a'], label="포트폴리오 A", lw=3); axs[1].plot(res['drawdown_a'], label="A 낙폭"); axs[1].fill_between(res['drawdown_a'].index, res['drawdown_a'], alpha=0.15)
        if res['asset_values_b'] is not None: axs[0].plot(res['asset_values_b'], label="포트폴리오 B", lw=3); axs[1].plot(res['drawdown_b'], label="B 낙폭")
        axs[0].plot(res['asset_values_bench'], label=f"벤치마크({bench})", ls="--", color="gray", alpha=0.7); axs[0].set_title(f"누적 자산 성장 추이 ({curr})", fontsize=16, pad=30, fontweight='bold'); axs[0].legend(fontsize=12); axs[0].grid(True, ls=":", alpha=0.6)
        from matplotlib.ticker import FuncFormatter, MaxNLocator
        axs[0].yaxis.set_major_formatter(FuncFormatter(lambda x, p: format(int(x), ','))); axs[0].yaxis.set_major_locator(MaxNLocator(nbins=12)); axs[1].plot(res['drawdown_bench'], label="벤치마크 낙폭", ls="--"); axs[1].legend(fontsize=12); axs[1].grid(True, ls=":", alpha=0.6); axs[1].set_title("최대 낙폭 (Drawdown) 분석 (%)", fontsize=16, pad=30, fontweight='bold'); fig.tight_layout(pad=8.0); FigureCanvasTkAgg(fig, master=t2_inner).get_tk_widget().pack(fill="x", pady=20, anchor="n")
        t3_inner = create_scroll_tab(nb, "3. 자산 상관관계"); fig_c = Figure(figsize=UNIFIED_FIG); axs_c = fig_c.subplots(2, 1)
        def draw_corr(ax, df, title):
            if not df.empty:
                im = ax.imshow(df.values, cmap='RdYlGn', vmin=-1, vmax=1, aspect='auto'); ax.set_title(title, pad=35, fontsize=16, fontweight='bold'); ax.set_xticks(range(len(df.columns))); ax.set_yticks(range(len(df.index))); ax.set_xticklabels(df.columns, rotation=45, ha='right', fontsize=11); ax.set_yticklabels(df.index, fontsize=11)
                for i in range(len(df.index)):
                    for j in range(len(df.columns)): ax.text(j, i, f"{df.iloc[i,j]:.2f}", ha="center", va="center", fontsize=11, fontweight='bold')
                fig_c.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
            else: ax.text(0.5, 0.5, "구성 종목이 부족하거나 데이터가 없습니다.", ha='center', va='center')
        draw_corr(axs_c[0], res['corr_a'], "포트폴리오 A 자산 상관관계"); draw_corr(axs_c[1], res['corr_b'], "포트폴리오 B 자산 상관관계"); fig_c.tight_layout(pad=10.0); FigureCanvasTkAgg(fig_c, master=t3_inner).get_tk_widget().pack(fill="x", pady=20, anchor="n")
        t4_inner = create_scroll_tab(nb, "4. 월간 수익률"); fig_m = Figure(figsize=UNIFIED_FIG); axs_m = fig_m.subplots(2, 1)
        def draw_m(ax, df, title):
            if not df.empty:
                im = ax.imshow(df.values, cmap='RdYlGn', vmin=-5, vmax=5, aspect='auto'); ax.set_title(title, pad=35, fontsize=16, fontweight='bold'); ax.set_yticks(range(len(df.index))); ax.set_yticklabels(df.index, fontsize=12); ax.set_xticks(range(len(df.columns))); ax.set_xticklabels([f"{m}월" for m in df.columns], fontsize=12)
                for i in range(len(df.index)):
                    for j in range(len(df.columns)):
                        val = df.iloc[i, j]
                        if pd.notna(val): txt = ax.annotate(f"{float(val):.1f}", xy=(j, i), ha="center", va="center", color="black", fontsize=11, fontweight='bold'); txt.set_path_effects([patheffects.withStroke(linewidth=2, foreground='white')])
                fig_m.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
            else: ax.text(0.5, 0.5, "데이터가 없습니다.", ha='center', va='center')
        draw_m(axs_m[0], res['monthly_a'], "포트폴리오 A 월간 수익률 (%)"); draw_m(axs_m[1], res['monthly_b'], "포트폴리오 B 월간 수익률 (%)"); fig_m.tight_layout(pad=10.0); FigureCanvasTkAgg(fig_m, master=t4_inner).get_tk_widget().pack(fill="x", pady=20, anchor="n")
        t5_inner = create_scroll_tab(nb, "5. 수익 기여도"); fig_con = Figure(figsize=UNIFIED_FIG); axs_con = fig_con.subplots(2, 1)
        def draw_contribution(ax, contrib_series, title):
            if contrib_series is not None and not contrib_series.empty and contrib_series.sum() != 0:
                c = contrib_series[contrib_series != 0]; c = c.sort_values(); names = [backtest_proxy.get_asset_name(t).split(' (')[0] for t in c.index]; colors = ['#27AE60' if x > 0 else '#E74C3C' for x in c.values]; bars = ax.barh(names, c.values * 100, color=colors, alpha=0.8); ax.set_title(title, pad=30, fontsize=16, fontweight='bold'); ax.set_xlabel("전체 수익에 대한 기여도 (%)", fontsize=12); ax.grid(True, axis='x', ls=':', alpha=0.6)
                for bar in bars: width = bar.get_width(); ax.text(width, bar.get_y() + bar.get_height()/2, f'{width:.1f}%', va='center', ha='left' if width > 0 else 'right', fontsize=10, fontweight='bold')
            else: ax.text(0.5, 0.5, "데이터가 없습니다.", ha='center', va='center')
        draw_contribution(axs_con[0], res['contrib_a'], "포트폴리오 A 자산별 성과 기여도"); draw_contribution(axs_con[1], res['contrib_b'], "포트폴리오 B 자산별 성과 기여도"); fig_con.tight_layout(pad=10.0); FigureCanvasTkAgg(fig_con, master=t5_inner).get_tk_widget().pack(fill="x", pady=20, anchor="n")

    def _refresh_sub_options(self):
        subs = sorted(self.df_master['sub'].dropna().unique().tolist())
        self.sub_cb["values"] = ["전체"] + subs; self.sub_var.set("전체"); self._refresh_group_options()
    def _refresh_group_options(self):
        s = self.sub_var.get(); df = self.df_master
        if s != "전체": df = df[df['sub'] == s]
        groups = sorted(df['group'].dropna().unique().tolist()); self.group_cb["values"] = ["전체"] + groups; self.group_var.set("전체")
    def _on_sub_changed(self, event): self._refresh_group_options(); self._refresh_asset_tree()
    def _on_header_click(self, col):
        if self.sort_col == col: self.sort_asc = not self.sort_asc
        else: self.sort_col = col; self.sort_asc = True
        self._refresh_asset_tree()
    def _refresh_asset_tree(self):
        for i in self.asset_tree.get_children(): self.asset_tree.delete(i)
        s, g, search = self.sub_var.get(), self.group_var.get(), self.search_var.get().strip().lower(); df = self.df_master.copy()
        if s != "전체": df = df[df['sub'] == s]
        if g != "전체": df = df[df['group'] == g]
        if search: df = df[df['ticker'].astype(str).str.lower().str.contains(search, na=False) | df['name'].astype(str).str.lower().str.contains(search, na=False)]
        if self.sort_col in df.columns: df = df.sort_values(by=self.sort_col, ascending=self.sort_asc)
        for _, row in df.head(500).iterrows():
            listed_val = str(row['listed']).split(' ')[0] if pd.notna(row['listed']) else "-"
            self.asset_tree.insert("", "end", values=(row['ticker'], row['name'], listed_val))

if __name__ == "__main__": 
    app = ProBacktestGUI(); app.mainloop()
