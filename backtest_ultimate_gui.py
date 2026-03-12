import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
import json
import io
import os
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import backtest_proxy
import backtest_dynamic
import backtest_hybrid
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

class UltimateBacktestGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("K-Global Asset Allocation Ultimate Edition v1.1")
        self.geometry("1550x980")
        self.configure(bg=COLOR_BG)

        # 1. 공통 상태 변수
        self.start_var = tk.StringVar(value="2005-01-01")
        self.initial_var = tk.StringVar(value="300000000")
        self.monthly_contrib_var = tk.StringVar(value="0")
        self.bench_var = tk.StringVar(value="SPY")
        self.currency_var = tk.StringVar(value="KRW")
        self.rebalance_var = tk.StringVar(value="Monthly")
        
        # 2. 정적 모드 변수
        self.port_a = {} 
        self.port_b = {}
        
        # 3. 동적 모드 변수
        self.dyn_strategy_type = tk.StringVar(value="VAA")
        self.dyn_off_universe = []
        self.dyn_def_universe = ["BIL", "SHY", "IEF"]
        self.dyn_canary_universe = ["VWO", "BND"]
        self.dyn_top_n_var = tk.StringVar(value="1")
        
        # 4. 하이브리드 모드 전용 변수 (독립 설정용)
        self.hybrid_port_a = {}
        self.hybrid_dyn_strategy_type = tk.StringVar(value="VAA")
        self.hybrid_dyn_off_universe = []
        self.hybrid_dyn_def_universe = ["BIL", "SHY", "IEF"]
        self.hybrid_dyn_canary_universe = ["VWO", "BND"]
        self.hybrid_dyn_top_n_var = tk.StringVar(value="1")
        self.sleeve_static_var = tk.StringVar(value="0.5")
        self.sleeve_dynamic_var = tk.StringVar(value="0.5")

        self.drag_data = {"source": None, "ticker": None, "port": None, "iid": None}

        self._apply_styles()
        self._build_ui()

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
        root = ttk.Frame(self, padding=20)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=3) 
        root.columnconfigure(1, weight=7) 

        # --- SIDEBAR ---
        sidebar = ttk.Frame(root)
        sidebar.grid(row=0, column=0, sticky="nsew", padx=(0, 15))

        # Global Config
        c_conf = ttk.Frame(sidebar, style="Card.TFrame", padding=15)
        c_conf.pack(fill="x", pady=(0, 15))
        ttk.Label(c_conf, text="Global Configuration", style="Header.TLabel").pack(anchor="w", pady=(0, 10))
        grid_f = ttk.Frame(c_conf, style="Card.TFrame"); grid_f.pack(fill="x")
        labels = [("Start Date", self.start_var), ("Initial Inv.", self.initial_var), ("Monthly Contrib.", self.monthly_contrib_var), ("Benchmark", self.bench_var), ("Currency", self.currency_var)]
        for i, (txt, var) in enumerate(labels):
            ttk.Label(grid_f, text=txt, style="CardLabel.TLabel").grid(row=i, column=0, sticky="w", pady=5)
            if txt == "Benchmark":
                ttk.Combobox(grid_f, textvariable=var, values=["SPY", "QQQ", "EWY", "VTI", "069500.KS"], width=13).grid(row=i, column=1, sticky="e")
            elif txt == "Currency":
                ttk.Combobox(grid_f, textvariable=var, values=["KRW", "USD"], state="readonly", width=13).grid(row=i, column=1, sticky="e")
            else:
                ttk.Entry(grid_f, textvariable=var, width=15).grid(row=i, column=1, sticky="e")

        # Asset Universe
        c_univ = ttk.Frame(sidebar, style="Card.TFrame", padding=15)
        c_univ.pack(fill="both", expand=True)
        ttk.Label(c_univ, text="Asset Universe (Drag to Tabs)", style="Header.TLabel").pack(anchor="w", pady=(0, 10))
        self.asset_tree = ttk.Treeview(c_univ, columns=("ticker", "name"), show="tree headings", height=15)
        self.asset_tree.heading("#0", text="Category"); self.asset_tree.heading("ticker", text="Ticker"); self.asset_tree.heading("name", text="Asset Name")
        self.asset_tree.column("#0", width=100); self.asset_tree.column("ticker", width=60); self.asset_tree.column("name", width=140)
        self.asset_tree.pack(fill="both", expand=True)
        self.asset_tree.bind("<Button-1>", self._start_universe_drag); self.asset_tree.bind("<B1-Motion>", self._drag_motion); self.asset_tree.bind("<ButtonRelease-1>", self._drop)
        for cat, assets in backtest_proxy.ASSET_UNIVERSE.items():
            parent = self.asset_tree.insert("", "end", text=cat, open=(cat=="주식 (Stocks)"))
            for t, n in assets.items(): self.asset_tree.insert(parent, "end", values=(t, n.split(' (')[0]))

        # --- MAIN AREA ---
        main_area = ttk.Frame(root)
        main_area.grid(row=0, column=1, sticky="nsew")
        self.mode_tabs = ttk.Notebook(main_area)
        self.mode_tabs.pack(fill="both", expand=True)

        self.tab_static = ttk.Frame(self.mode_tabs, padding=15)
        self.mode_tabs.add(self.tab_static, text=" Strategic Allocation ")
        self._build_static_tab()

        self.tab_dynamic = ttk.Frame(self.mode_tabs, padding=15)
        self.mode_tabs.add(self.tab_dynamic, text=" Dynamic Strategy ")
        self._build_dynamic_tab()

        self.tab_hybrid = ttk.Frame(self.mode_tabs, padding=15)
        self.mode_tabs.add(self.tab_hybrid, text=" Hybrid Engine ")
        self._build_hybrid_tab()

    def _build_static_tab(self):
        top_bar = ttk.Frame(self.tab_static, padding=(0, 0, 0, 15))
        top_bar.pack(fill="x")
        ttk.Label(top_bar, text="Strategic Portfolio Design", style="Header.TLabel").pack(side="left")
        
        preset_f = ttk.Frame(top_bar); preset_f.pack(side="left", padx=30)
        ttk.Label(preset_f, text="Presets:", style="CardLabel.TLabel").pack(side="left", padx=5)
        self.preset_cb = ttk.Combobox(preset_f, values=["선택 안함"] + list(backtest_proxy.STRATEGY_PRESETS.keys()), state="readonly", width=20)
        self.preset_cb.set("선택 안함"); self.preset_cb.pack(side="left", padx=5)
        ttk.Button(preset_f, text="Apply to A", command=lambda: self._apply_preset("A")).pack(side="left", padx=2)
        ttk.Button(preset_f, text="Apply to B", command=lambda: self._apply_preset("B")).pack(side="left", padx=2)

        ttk.Button(top_bar, text="💾 Save", command=self._save_config).pack(side="right", padx=5)
        ttk.Button(top_bar, text="📂 Load", command=self._load_config).pack(side="right")

        ports_f = ttk.Frame(self.tab_static); ports_f.pack(fill="both", expand=True)
        ports_f.columnconfigure(0, weight=1); ports_f.columnconfigure(1, weight=1)
        self.box_a = self._create_port_box(ports_f, "Portfolio A", "A")
        self.box_a.grid(row=0, column=0, sticky="nsew", padx=10)
        self.box_b = self._create_port_box(ports_f, "Portfolio B", "B")
        self.box_b.grid(row=0, column=1, sticky="nsew", padx=10)
        
        run_f = ttk.Frame(self.tab_static, padding=(0, 15, 0, 0))
        run_f.pack(fill="x")
        ttk.Label(run_f, text="Rebalancing:", style="CardLabel.TLabel").pack(side="left", padx=5)
        ttk.Combobox(run_f, textvariable=self.rebalance_var, values=["Monthly", "Quarterly", "Yearly", "None"], state="readonly", width=10).pack(side="left", padx=5)
        ttk.Button(run_f, text="🚀 RUN STRATEGIC ANALYSIS", style="Success.TButton", command=self._run_static).pack(side="right", fill="x", expand=True, ipady=10)

    def _build_dynamic_tab(self):
        main_f = ttk.Frame(self.tab_dynamic, style="Card.TFrame", padding=20)
        main_f.pack(fill="both", expand=True)
        ttk.Label(main_f, text="Dynamic Algorithm Setup", style="Header.TLabel").pack(anchor="w", pady=(0, 20))
        
        conf_f = ttk.Frame(main_f, style="Card.TFrame"); conf_f.pack(fill="x", pady=10)
        ttk.Label(conf_f, text="Strategy:", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w", padx=5)
        self.dyn_cb = ttk.Combobox(conf_f, textvariable=self.dyn_strategy_type, values=["VAA", "DAA", "GEM"], state="readonly", width=15)
        self.dyn_cb.grid(row=0, column=1, sticky="w", padx=10); self.dyn_cb.bind("<<ComboboxSelected>>", self._on_dyn_strategy_change)
        ttk.Label(conf_f, text="Top N:", style="CardLabel.TLabel").grid(row=0, column=2, sticky="w", padx=20)
        ttk.Entry(conf_f, textvariable=self.dyn_top_n_var, width=5).grid(row=0, column=3, sticky="w")

        univ_f = ttk.Frame(main_f, style="Card.TFrame", padding=(0, 20)); univ_f.pack(fill="both", expand=True); univ_f.columnconfigure(0, weight=1); univ_f.columnconfigure(1, weight=1)
        def build_dyn_list(parent, title, col):
            f = ttk.LabelFrame(parent, text=title, padding=10); f.grid(row=0, column=col, sticky="nsew", padx=10)
            lb = tk.Listbox(f, font=FONT_MAIN, height=10, bg="white", borderwidth=0, highlightthickness=1, highlightcolor=COLOR_BORDER); lb.pack(fill="both", expand=True); return lb, f
        self.dyn_off_lb, self.dyn_off_box = build_dyn_list(univ_f, "Offensive Assets (Drag Here)", 0)
        self.dyn_def_lb, self.dyn_def_box = build_dyn_list(univ_f, "Defensive Assets (Drag Here)", 1)
        btn_f = ttk.Frame(main_f, style="Card.TFrame"); btn_f.pack(fill="x", pady=20)
        ttk.Button(btn_f, text="Clear Dynamic Universe", command=self._clear_dyn_univ).pack(side="left")
        ttk.Button(main_f, text="🚀 RUN DYNAMIC ANALYSIS", style="Success.TButton", command=self._run_dynamic).pack(fill="x", ipady=12)

    def _build_hybrid_tab(self):
        main_scroll = tk.Canvas(self.tab_hybrid, bg=COLOR_BG, highlightthickness=0)
        scroll_y = ttk.Scrollbar(self.tab_hybrid, orient="vertical", command=main_scroll.yview)
        container = ttk.Frame(main_scroll, padding=10)
        
        container.bind("<Configure>", lambda e: main_scroll.configure(scrollregion=main_scroll.bbox("all")))
        main_scroll.create_window((0, 0), window=container, anchor="nw", width=1250)
        main_scroll.configure(yscrollcommand=scroll_y.set)
        
        main_scroll.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

        ttk.Label(container, text="Hybrid Engine: Complete Setup", style="Header.TLabel").pack(anchor="w", pady=(0, 20))

        # --- Main Split Content ---
        top_split = ttk.Frame(container); top_split.pack(fill="x")
        top_split.columnconfigure(0, weight=1); top_split.columnconfigure(1, weight=1)

        # --- LEFT: Strategic Config ---
        left_f = ttk.Frame(top_split, padding=10); left_f.grid(row=0, column=0, sticky="nsew")
        l_conf_f = ttk.Frame(left_f); l_conf_f.pack(fill="x", pady=(0, 10))
        ttk.Label(l_conf_f, text="Preset:", font=FONT_BOLD).pack(side="left", padx=5)
        self.hybrid_preset_cb = ttk.Combobox(l_conf_f, values=["선택 안함"] + list(backtest_proxy.STRATEGY_PRESETS.keys()), state="readonly", width=18)
        self.hybrid_preset_cb.set("선택 안함"); self.hybrid_preset_cb.pack(side="left", padx=5)
        ttk.Button(l_conf_f, text="Apply", command=lambda: self._apply_preset("Hybrid_A"), width=10).pack(side="left")

        self.box_hybrid_a = self._create_port_box(left_f, "Strategic Sleeve (Port A)", "Hybrid_A")
        self.box_hybrid_a.pack(fill="both", expand=True)

        # --- RIGHT: Dynamic Config ---
        right_f = ttk.Frame(top_split, padding=10); right_f.grid(row=0, column=1, sticky="nsew")
        r_conf_f = ttk.Frame(right_f); r_conf_f.pack(fill="x", pady=(0, 10))
        ttk.Label(r_conf_f, text="Strategy:", font=FONT_BOLD).pack(side="left", padx=5)
        self.hybrid_dyn_cb = ttk.Combobox(r_conf_f, textvariable=self.hybrid_dyn_strategy_type, values=["VAA", "DAA", "GEM"], state="readonly", width=12)
        self.hybrid_dyn_cb.pack(side="left", padx=5); self.hybrid_dyn_cb.bind("<<ComboboxSelected>>", lambda e: self._on_dyn_strategy_change(e, True))
        ttk.Label(r_conf_f, text="Top N:", font=FONT_BOLD).pack(side="left", padx=(15, 5))
        ttk.Entry(r_conf_f, textvariable=self.hybrid_dyn_top_n_var, width=5).pack(side="left")

        # Right Asset Box - Styled exactly like left port box
        dyn_asset_f = ttk.Frame(right_f, style="Card.TFrame", padding=15); dyn_asset_f.pack(fill="both", expand=True)
        ttk.Label(dyn_asset_f, text="Dynamic Sleeve Strategy", style="Header.TLabel").pack(anchor="w", pady=(0, 10))
        
        univ_f = ttk.Frame(dyn_asset_f, style="Card.TFrame"); univ_f.pack(fill="both", expand=True, pady=10)
        def build_list(parent, title):
            f = ttk.Frame(parent, style="Card.TFrame"); f.pack(fill="x", pady=2)
            ttk.Label(f, text=title, font=FONT_BOLD, style="CardLabel.TLabel").pack(anchor="w")
            lb = tk.Listbox(f, font=FONT_MAIN, height=8, bg="white", borderwidth=1, relief="solid"); lb.pack(fill="x"); return lb, f
        self.hybrid_dyn_off_lb, self.hybrid_dyn_off_box = build_list(univ_f, "Offensive Assets (Drag Here)")
        self.hybrid_dyn_def_lb, self.hybrid_dyn_def_box = build_list(univ_f, "Defensive Assets (Drag Here)")
        
        btn_box = ttk.Frame(dyn_asset_f, style="Card.TFrame"); btn_box.pack(fill="x", pady=(10, 0))
        ttk.Button(btn_box, text="Clear", width=8, command=lambda: self._clear_dyn_univ(True)).pack(side="left")

        # --- Bottom Section: Blend & Run ---
        blend_card = ttk.Frame(container, style="Card.TFrame", padding=20); blend_card.pack(fill="x", pady=20)
        ttk.Label(blend_card, text="Final Portfolio Blend (%)", style="Header.TLabel").pack(pady=(0, 15))
        
        # Use tk.Frame for direct background support or keep ttk.Frame with style
        b_grid = ttk.Frame(blend_card, style="Card.TFrame"); b_grid.pack()
        ttk.Label(b_grid, text="Static Sleeve", font=FONT_BOLD, style="CardLabel.TLabel").grid(row=0, column=0, padx=20)
        ttk.Entry(b_grid, textvariable=self.sleeve_static_var, width=10, justify="center").grid(row=1, column=0)
        # Use standard tk.Label for the "+" sign to allow direct background color
        tk.Label(b_grid, text="+", font=("Arial", 20), bg=COLOR_CARD, fg=COLOR_PRIMARY).grid(row=1, column=1)
        ttk.Label(b_grid, text="Dynamic Sleeve", font=FONT_BOLD, style="CardLabel.TLabel").grid(row=0, column=2, padx=20)
        ttk.Entry(b_grid, textvariable=self.sleeve_dynamic_var, width=10, justify="center").grid(row=1, column=2)

        ttk.Button(container, text="🚀 RUN INTEGRATED HYBRID BACKTEST", style="Success.TButton", command=self._run_hybrid).pack(fill="x", ipady=15, pady=10)

    def _create_port_box(self, parent, title, which):
        frame = ttk.Frame(parent, style="Card.TFrame", padding=15)
        ttk.Label(frame, text=title, style="Header.TLabel").pack(anchor="w", pady=(0, 10))
        tree = ttk.Treeview(frame, columns=("ticker", "name", "weight"), show="headings", height=8)
        tree.heading("ticker", text="Ticker"); tree.heading("name", text="Name"); tree.heading("weight", text="W(%)")
        tree.column("ticker", width=70, anchor="center"); tree.column("name", width=180); tree.column("weight", width=60, anchor="center")
        tree.pack(fill="x"); tree.bind("<Button-1>", self._start_port_drag); tree.bind("<B1-Motion>", self._drag_motion); tree.bind("<ButtonRelease-1>", self._drop)
        edit_f = ttk.Frame(frame, style="Card.TFrame"); edit_f.pack(fill="both", expand=True, pady=(10, 0))
        canvas = tk.Canvas(edit_f, bg="white", highlightthickness=0, height=200); scrollbar = ttk.Scrollbar(edit_f, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas, style="Card.TFrame"); inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw"); canvas.configure(yscrollcommand=scrollbar.set); canvas.pack(side="left", fill="both", expand=True); scrollbar.pack(side="right", fill="y")
        for w in [frame, tree, edit_f, canvas, inner]: w.bind("<ButtonRelease-1>", self._drop, add="+")
        btn_box = ttk.Frame(frame, style="Card.TFrame"); btn_box.pack(fill="x", pady=(10, 0))
        ttk.Button(btn_box, text="Del", width=5, command=lambda: self._delete_asset(which)).pack(side="left")
        ttk.Button(btn_box, text="Equal", width=8, command=lambda: self._equal_weight(which)).pack(side="left", padx=5)
        sum_lbl = ttk.Label(btn_box, text="Sum: 0%", style="CardLabel.TLabel"); sum_lbl.pack(side="right")
        
        if which == "A": self.tree_a, self.edit_a, self.sum_a, self.box_a = tree, inner, sum_lbl, frame
        elif which == "B": self.tree_b, self.edit_b, self.sum_b, self.box_b = tree, inner, sum_lbl, frame
        elif which == "Hybrid_A": self.tree_h_a, self.edit_h_a, self.sum_h_a, self.box_h_a = tree, inner, sum_lbl, frame
        return frame

    def _run_static(self):
        try:
            if not self.port_a or not self.port_b: raise ValueError("Port A/B 모두 설정 필요")
            res = backtest_proxy.run_pro_backtest(self.port_a, self.port_b, start=self.start_var.get(), initial_investment=int(self.initial_var.get()), benchmark_ticker=self.bench_var.get(), rebalance=self.rebalance_var.get(), base_currency=self.currency_var.get(), monthly_contribution=float(self.monthly_contrib_var.get()))
            self._show_results(res, "Strategic Analysis")
        except Exception as e: messagebox.showerror("Error", str(e))

    def _run_dynamic(self):
        try:
            if not self.dyn_off_universe: raise ValueError("공격 자산군 설정 필요")
            stype = self.dyn_strategy_type.get()
            res = backtest_dynamic.run_dynamic_strategy(stype, self.dyn_off_universe, self.dyn_def_universe, canary_universe=self.dyn_canary_universe, start=self.start_var.get(), initial_investment=int(self.initial_var.get()), top_n=int(self.dyn_top_n_var.get()), base_currency=self.currency_var.get(), monthly_contribution=float(self.monthly_contrib_var.get()), benchmark_ticker=self.bench_var.get())
            res.update({"is_dynamic": True, "title": f"Dynamic: {stype}", "strategy_type": stype})
            self._show_results(res, res["title"])
        except Exception as e: messagebox.showerror("Error", str(e))

    def _run_hybrid(self):
        try:
            stype = self.hybrid_dyn_strategy_type.get()
            blend = {"static": float(self.sleeve_static_var.get()), "dynamic": float(self.sleeve_dynamic_var.get())}
            dyn_cfg = {"type": stype, "off": self.hybrid_dyn_off_universe, "def": self.hybrid_dyn_def_universe, "canary": self.hybrid_dyn_canary_universe, "top_n": int(self.hybrid_dyn_top_n_var.get())}
            res = backtest_hybrid.run_hybrid_backtest(self.hybrid_port_a, dyn_cfg, blend, start=self.start_var.get(), initial_investment=int(self.initial_var.get()), benchmark_ticker=self.bench_var.get(), base_currency=self.currency_var.get(), monthly_contribution=float(self.monthly_contrib_var.get()))
            res.update({"is_hybrid": True, "title": "Hybrid Combination", "strategy_type": stype})
            self._show_results(res, res["title"])
        except Exception as e: messagebox.showerror("Error", str(e))

    def _show_results(self, res, title):
        win = tk.Toplevel(self); win.title(title); win.geometry("1350x950"); win.configure(bg=COLOR_BG)
        top = ttk.Frame(win, padding=(20, 10)); top.pack(fill="x")
        curr = self.currency_var.get(); ttk.Label(top, text=f"{title} ({curr} Base)", font=FONT_TITLE).pack(side="left")
        
        def export_excel():
            path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if path:
                try:
                    with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
                        res['metrics'].to_excel(writer, sheet_name='Summary')
                        if 'weights' in res: res['weights'].to_excel(writer, sheet_name='Holding_History')
                        if 'monthly_a' in res: res['monthly_a'].to_excel(writer, sheet_name='Monthly_Returns')
                    messagebox.showinfo("Success", "Professional Report Saved.")
                except Exception as e: messagebox.showerror("Error", str(e))
        ttk.Button(top, text="📥 Export Excel", style="Success.TButton", command=export_excel).pack(side="right")
        
        nb = ttk.Notebook(win); nb.pack(fill="both", expand=True, padx=20, pady=10)
        def create_scroll_tab(notebook, t_text):
            frame = ttk.Frame(notebook); notebook.add(frame, text=t_text); canvas = tk.Canvas(frame, bg=COLOR_BG, highlightthickness=0); scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview); inner = ttk.Frame(canvas, style="Card.TFrame")
            inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
            def _on_canvas_configure(event): canvas.coords(canvas_window, event.width / 2, 0)
            canvas.bind("<Configure>", _on_canvas_configure); canvas_window = canvas.create_window((650, 0), window=inner, anchor="n"); canvas.configure(yscrollcommand=scrollbar.set); canvas.pack(side="left", fill="both", expand=True); scrollbar.pack(side="right", fill="y")
            def _on_mousewheel(event): canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel)); canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>")); return inner

        UNIFIED_FIG = (15, 12)
        # Tab 1: Stats
        t1 = create_scroll_tab(nb, "Analysis Stats")
        tree = ttk.Treeview(t1, columns=("M", "S", "C", "B"), show="headings", height=15)
        for c, h in zip(tree["columns"], ["Metric", "Strategy", "Comparison", "Benchmark"]): tree.heading(c, text=h); tree.column(c, anchor="center", width=230)
        tree.pack(pady=30); m = res['metrics']
        for idx, row in m.iterrows(): tree.insert("", "end", values=(idx, *row.values.tolist()))

        # Tab 2: Growth
        t2 = create_scroll_tab(nb, "Equity Growth"); fig = Figure(figsize=UNIFIED_FIG); axs = fig.subplots(2, 1)
        v_a = res.get('asset_values_a', res.get('asset_values'))
        axs[0].plot(v_a, label="Strategy", lw=3, color='#3498DB')
        if 'asset_values_b' in res and not res.get('is_dynamic'): axs[0].plot(res['asset_values_b'], label="Comparison", lw=3, color='#E67E22')
        if 'asset_values_bench' in res: axs[0].plot(res['asset_values_bench'], label="Benchmark", ls="--", color="gray", alpha=0.7)
        axs[0].set_title("Cumulative Growth", fontsize=16, fontweight='bold'); axs[0].legend(); axs[0].grid(True, ls=":")
        axs[0].yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format(int(x), ',')))
        dd_a = res.get('drawdown_a', res.get('drawdown'))
        axs[1].fill_between(dd_a.index, dd_a, color='#3498DB', alpha=0.15); axs[1].plot(dd_a, label="Strategy DD"); axs[1].set_title("Drawdown (%)", fontsize=16, fontweight='bold'); axs[1].grid(True, ls=":")
        fig.tight_layout(pad=5.0); FigureCanvasTkAgg(fig, master=t2).get_tk_widget().pack(fill="x", pady=20)

        # Tab 3: Holding History (Heatmap Style)
        if 'weights' in res:
            t3 = create_scroll_tab(nb, "Holding History")
            stype = res.get("strategy_type", "Unknown")
            info_f = ttk.LabelFrame(t3, text=f" Allocation Logic: {stype} ", padding=15)
            info_f.pack(fill="x", pady=(10, 20))
            logic_text = ""
            if stype == "VAA": logic_text = "💡 [VAA: 공격적 자산배분]\n공격 자산군 중 하나라도 모멘텀이 음수(-)로 꺾이면 위험 신호로 간주합니다.\n위험 신호 발생 시 즉시 가장 안전한 방어 자산으로 전량(100%) 대피하여 MDD를 방어합니다."
            elif stype == "DAA": logic_text = "💡 [DAA: 방어적 자산배분]\n카나리아 자산(VWO, BND)을 통해 시장의 하락 전조를 감시합니다.\n카나리아 자산의 위험 신호 개수에 따라 공격 자산 비중을 단계적(100% -> 50% -> 0%)으로 조절하는 유연한 전략입니다."
            elif stype == "GEM": logic_text = "💡 [GEM: 글로벌 주식 모멘텀]\n미국 주식과 세계 주식의 수익률을 비교하여 더 강한 시장에 투자합니다.\n두 시장 모두 모멘텀이 낮을 경우 안전 자산(채권)으로 이동하여 하락장을 피합니다."
            if res.get("is_hybrid"): logic_text += "\n\n⚙️ [Hybrid Mode] 본 시뮬레이션은 정적 자산배분(Port A)과 위 동적 전략을 혼합하여 운용합니다."
            ttk.Label(info_f, text=logic_text, font=FONT_MAIN, justify="left", wraplength=1200).pack(anchor="w")

            fig_h = Figure(figsize=(15, 10)); ax_h = fig_h.add_subplot(111)
            w_df = res['weights'].dropna(axis=1, how='all').T
            im = ax_h.imshow(w_df.values, cmap='Blues', aspect='auto', interpolation='nearest', vmin=0, vmax=1)
            ax_h.set_yticks(np.arange(len(w_df.index))); ax_h.set_yticklabels(w_df.index, fontsize=10, fontweight='bold')
            dates = w_df.columns; n_xticks = min(15, len(dates)); indices = np.linspace(0, len(dates)-1, n_xticks, dtype=int)
            ax_h.set_xticks(indices); ax_h.set_xticklabels([dates[i].strftime('%Y-%m') for i in indices], rotation=30, ha='right')
            ax_h.set_title("Asset Allocation History Heatmap", fontsize=16, fontweight='bold', pad=25)
            cbar = fig_h.colorbar(im, ax=ax_h, pad=0.02, aspect=30); cbar.set_label("Weight (%)", fontsize=11)
            cbar.ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x*100:.0f}%'))
            ax_h.set_yticks(np.arange(-.5, len(w_df.index), 1), minor=True); ax_h.grid(which='minor', color='white', linestyle='-', linewidth=2)
            fig_h.tight_layout(); FigureCanvasTkAgg(fig_h, master=t3).get_tk_widget().pack(fill="x", pady=20)

        if 'monthly_a' in res:
            t4 = create_scroll_tab(nb, "Monthly Matrix"); fig_m = Figure(figsize=(15, 10)); ax_m = fig_m.add_subplot(111)
            m_data = res['monthly_a']
            im_m = ax_m.imshow(m_data.values, cmap='RdYlGn', vmin=-5, vmax=5, aspect='auto')
            ax_m.set_title("Monthly Returns Matrix (%)", fontsize=15, fontweight='bold')
            
            # --- Add Return Values as Text ---
            for i in range(len(m_data.index)):
                for j in range(len(m_data.columns)):
                    val = m_data.iloc[i, j]
                    if not np.isnan(val):
                        # Contrast text color based on cell color intensity
                        color = "white" if abs(val) > 3 else "black"
                        ax_m.text(j, i, f"{val:.1f}", ha="center", va="center", color=color, fontsize=9, fontweight='bold')

            # Set Axis Labels
            ax_m.set_xticks(np.arange(len(m_data.columns))); ax_m.set_xticklabels(m_data.columns)
            ax_m.set_yticks(np.arange(len(m_data.index))); ax_m.set_yticklabels(m_data.index)
            
            fig_m.colorbar(im_m, ax=ax_m)
            FigureCanvasTkAgg(fig_m, master=t4).get_tk_widget().pack(fill="x", pady=20)

    def _refresh_port_ui(self, which):
        if which == "A": tree, edit, data, sum_lbl = self.tree_a, self.edit_a, self.port_a, self.sum_a
        elif which == "B": tree, edit, data, sum_lbl = self.tree_b, self.edit_b, self.port_b, self.sum_b
        elif which == "Hybrid_A": tree, edit, data, sum_lbl = self.tree_h_a, self.edit_h_a, self.hybrid_port_a, self.sum_h_a
        
        for i in tree.get_children(): tree.delete(i)
        for c in edit.winfo_children(): c.destroy()
        total_w = 0
        for t, w in data.items():
            name = backtest_proxy.get_asset_name(t).split(' (')[0]; tree.insert("", "end", values=(t, name, f"{w*100:.1f}")); total_w += w
            row = ttk.Frame(edit, style="Card.TFrame"); row.pack(fill="x", pady=2)
            ttk.Label(row, text=f"{t} | {name[:12]}", width=25, style="CardLabel.TLabel").pack(side="left")
            var = tk.StringVar(value=f"{w*100:.1f}"); ent = ttk.Entry(row, textvariable=var, width=8, justify="right"); ent.pack(side="right", padx=5)
            var.trace_add("write", lambda *a, tk=t, v=var, wh=which: self._update_weight(wh, tk, v))
        sum_lbl.config(text=f"Sum: {total_w*100:.1f}%")

    def _update_weight(self, which, ticker, var):
        try:
            val = float(var.get() or 0) / 100
            if which == "A": self.port_a[ticker] = val; lbl = self.sum_a; data = self.port_a
            elif which == "B": self.port_b[ticker] = val; lbl = self.sum_b; data = self.port_b
            elif which == "Hybrid_A": self.hybrid_port_a[ticker] = val; lbl = self.sum_h_a; data = self.hybrid_port_a
            lbl.config(text=f"Sum: {sum(data.values())*100:.1f}%")
        except: pass

    def _on_dyn_strategy_change(self, event, is_hybrid=False):
        stype = self.hybrid_dyn_strategy_type.get() if is_hybrid else self.dyn_strategy_type.get()
        self._clear_dyn_univ(is_hybrid)
        off, def_, canary = [], [], []
        if stype == "VAA": off, def_ = ["SPY", "VEA", "VWO", "AGG"], ["SHY", "BIL", "IEF"]
        elif stype == "DAA": off, def_, canary = ["SPY", "IWM", "VEA", "VWO", "VNQ", "GLD", "TLT", "LQD"], ["SHY", "BIL", "IEF"], ["VWO", "BND"]
        elif stype == "GEM": off, def_ = ["SPY", "VEA"], ["AGG"]
        if is_hybrid: self.hybrid_dyn_off_universe = off; self.hybrid_dyn_def_universe = def_; self.hybrid_dyn_canary_universe = canary
        else: self.dyn_off_universe = off; self.dyn_def_universe = def_; self.dyn_canary_universe = canary
        self._refresh_dyn_ui(is_hybrid)

    def _refresh_dyn_ui(self, is_hybrid=False):
        lb_off, lb_def, univ_off, univ_def = (self.hybrid_dyn_off_lb, self.hybrid_dyn_def_lb, self.hybrid_dyn_off_universe, self.hybrid_dyn_def_universe) if is_hybrid else (self.dyn_off_lb, self.dyn_def_lb, self.dyn_off_universe, self.dyn_def_universe)
        lb_off.delete(0, tk.END); [lb_off.insert(tk.END, f"{t} | {backtest_proxy.get_asset_name(t)}") for t in univ_off]
        lb_def.delete(0, tk.END); [lb_def.insert(tk.END, f"{t} | {backtest_proxy.get_asset_name(t)}") for t in univ_def]

    def _clear_dyn_univ(self, is_hybrid=False):
        if is_hybrid: self.hybrid_dyn_off_universe = []; self.hybrid_dyn_def_universe = []; self.hybrid_dyn_canary_universe = []
        else: self.dyn_off_universe = []; self.dyn_def_universe = []; self.dyn_canary_universe = []
        self._refresh_dyn_ui(is_hybrid)

    def _apply_preset(self, which):
        cb = self.hybrid_preset_cb if which == "Hybrid_A" else self.preset_cb
        p_name = cb.get()
        if p_name != "선택 안함":
            target = self.port_a if which == "A" else (self.port_b if which == "B" else self.hybrid_port_a)
            target.clear()
            for t, w in backtest_proxy.STRATEGY_PRESETS[p_name].items(): target[str(t)] = w
            self._refresh_port_ui(which)

    def _save_config(self):
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if path:
            try:
                config = {"start": self.start_var.get(), "initial": self.initial_var.get(), "contrib": self.monthly_contrib_var.get(), "bench": self.bench_var.get(), "cur": self.currency_var.get(), "reb": self.rebalance_var.get(), "port_a": self.port_a, "port_b": self.port_b}
                with open(path, 'w', encoding='utf-8') as f: json.dump(config, f, indent=4, ensure_ascii=False)
                messagebox.showinfo("Success", "Saved.")
            except Exception as e: messagebox.showerror("Error", str(e))

    def _load_config(self):
        path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if path:
            try:
                with open(path, 'r', encoding='utf-8') as f: cfg = json.load(f)
                self.start_var.set(cfg.get("start", self.start_var.get())); self.initial_var.set(cfg.get("initial", self.initial_var.get())); self.monthly_contrib_var.set(cfg.get("contrib", self.monthly_contrib_var.get()))
                self.bench_var.set(cfg.get("bench", self.bench_var.get())); self.currency_var.set(cfg.get("cur", self.currency_var.get())); self.rebalance_var.set(cfg.get("reb", self.rebalance_var.get()))
                self.port_a = cfg.get("port_a", cfg.get("a", {})); self.port_b = cfg.get("port_b", cfg.get("b", {})); self._refresh_port_ui("A"); self._refresh_port_ui("B"); messagebox.showinfo("Success", "Applied.")
            except Exception as e: messagebox.showerror("Error", str(e))

    def _delete_asset(self, which):
        if which == "A": tree, data = self.tree_a, self.port_a
        elif which == "B": tree, data = self.tree_b, self.port_b
        elif which == "Hybrid_A": tree, data = self.tree_h_a, self.hybrid_port_a
        for it in tree.selection():
            t = str(tree.item(it)['values'][0])
            if t in data: del data[t]
        self._refresh_port_ui(which)

    def _equal_weight(self, which):
        if which == "A": data = self.port_a
        elif which == "B": data = self.port_b
        elif which == "Hybrid_A": data = self.hybrid_port_a
        if data:
            eq = 1.0 / len(data)
            for t in data: data[t] = eq
            self._refresh_port_ui(which)

    def _start_universe_drag(self, event):
        iid = self.asset_tree.identify_row(event.y)
        if iid: vals = self.asset_tree.item(iid)['values']; self.drag_data = {"source": "universe", "ticker": str(vals[0]), "iid": iid}; self.asset_tree.config(cursor="hand2")
    def _start_port_drag(self, event):
        widget = event.widget; iid = widget.identify_row(event.y)
        if iid: self.drag_data = {"source": "portfolio", "ticker": str(widget.item(iid)['values'][0]), "port": "A" if widget == self.tree_a else ("B" if widget == self.tree_b else "Hybrid_A"), "iid": iid}; widget.config(cursor="hand2")
    def _drag_motion(self, event): pass
    def _drop(self, event):
        if not self.drag_data["ticker"]: return
        x, y = event.x_root, event.y_root; target = self.winfo_containing(x, y); ticker = self.drag_data["ticker"]
        def is_desc(c, p):
            while c:
                if c == p: return True
                c = c.master
            return False
        if is_desc(target, self.box_a): 
            if ticker not in self.port_a: self.port_a[ticker] = 0.0; self._refresh_port_ui("A")
        elif is_desc(target, self.box_b):
            if ticker not in self.port_b: self.port_b[ticker] = 0.0; self._refresh_port_ui("B")
        elif is_desc(target, self.box_h_a):
            if ticker not in self.hybrid_port_a: self.hybrid_port_a[ticker] = 0.0; self._refresh_port_ui("Hybrid_A")
        elif is_desc(target, self.dyn_off_box):
            if ticker not in self.dyn_off_universe: self.dyn_off_universe.append(ticker); self._refresh_dyn_ui()
        elif is_desc(target, self.dyn_def_box):
            if ticker not in self.dyn_def_universe: self.dyn_def_universe.append(ticker); self._refresh_dyn_ui()
        elif is_desc(target, self.hybrid_dyn_off_box):
            if ticker not in self.hybrid_dyn_off_universe: self.hybrid_dyn_off_universe.append(ticker); self._refresh_dyn_ui(True)
        elif is_desc(target, self.hybrid_dyn_def_box):
            if ticker not in self.hybrid_dyn_def_universe: self.hybrid_dyn_def_universe.append(ticker); self._refresh_dyn_ui(True)
        self.asset_tree.config(cursor=""); self.drag_data = {"ticker": None}

if __name__ == "__main__":
    app = UltimateBacktestGUI()
    app.mainloop()
