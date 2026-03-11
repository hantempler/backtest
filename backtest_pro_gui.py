import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
import json
import io
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
        self.title("K-Global 하이브리드 자산배분 시뮬레이터 v1.4")
        self.geometry("1400x950")
        self.configure(bg=COLOR_BG)

        self.port_a = {}
        self.port_b = {}
        self.start_var = tk.StringVar(value="2005-01-01")
        self.initial_var = tk.StringVar(value="300000000")
        self.bench_var = tk.StringVar(value="SPY")
        self.rebalance_var = tk.StringVar(value="Monthly")
        self.currency_var = tk.StringVar(value="KRW")
        self.custom_ticker_var = tk.StringVar()
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

        sidebar = ttk.Frame(root)
        sidebar.grid(row=0, column=0, sticky="nsew", padx=(0, 15))

        cp = ttk.Frame(sidebar, style="Card.TFrame", padding=15)
        cp.pack(fill="x", pady=(0, 15))
        ttk.Label(cp, text="Strategy Presets", style="Header.TLabel").pack(anchor="w", pady=(0, 10))
        self.preset_cb = ttk.Combobox(cp, values=["선택 안함"] + list(backtest_proxy.STRATEGY_PRESETS.keys()), state="readonly")
        self.preset_cb.set("선택 안함")
        self.preset_cb.pack(fill="x", pady=(0, 10))
        btn_pf = ttk.Frame(cp, style="Card.TFrame")
        btn_pf.pack(fill="x")
        ttk.Button(btn_pf, text="Apply to A", command=lambda: self._apply_preset("A")).pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(btn_pf, text="Apply to B", command=lambda: self._apply_preset("B")).pack(side="left", fill="x", expand=True)

        c1 = ttk.Frame(sidebar, style="Card.TFrame", padding=15)
        c1.pack(fill="both", expand=True, pady=(0, 15))
        ttk.Label(c1, text="Asset Universe (Drag to Ports)", style="Header.TLabel").pack(anchor="w", pady=(0, 10))
        self.asset_tree = ttk.Treeview(c1, columns=("ticker", "name"), show="tree headings", height=12)
        self.asset_tree.heading("#0", text="Category")
        self.asset_tree.heading("ticker", text="Ticker")
        self.asset_tree.heading("name", text="Asset Name")
        self.asset_tree.column("#0", width=120); self.asset_tree.column("ticker", width=60, anchor="center"); self.asset_tree.column("name", width=150)
        self.asset_tree.pack(fill="both", expand=True)
        self.asset_tree.bind("<Button-1>", self._start_universe_drag)
        self.asset_tree.bind("<B1-Motion>", self._drag_motion)
        self.asset_tree.bind("<ButtonRelease-1>", self._drop)
        for cat, assets in backtest_proxy.ASSET_UNIVERSE.items():
            parent = self.asset_tree.insert("", "end", text=cat, open=True)
            for t, n in assets.items(): self.asset_tree.insert(parent, "end", values=(t, n.split(' (')[0]))
        
        btn_f = ttk.Frame(c1, style="Card.TFrame")
        btn_f.pack(fill="x", pady=(10, 0))
        ttk.Button(btn_f, text="+ A", width=5, command=lambda: self._add_to_port("A")).pack(side="left", padx=(0, 5))
        ttk.Button(btn_f, text="+ B", width=5, command=lambda: self._add_to_port("B")).pack(side="left")
        ttk.Label(btn_f, text="Custom:", style="CardLabel.TLabel").pack(side="left", padx=(10, 5))
        ttk.Entry(btn_f, textvariable=self.custom_ticker_var, width=8).pack(side="left")
        ttk.Button(btn_f, text="Add", width=5, command=self._add_custom).pack(side="left", padx=5)

        c2 = ttk.Frame(sidebar, style="Card.TFrame", padding=15)
        c2.pack(fill="x")
        ttk.Label(c2, text="Configuration", style="Header.TLabel").pack(anchor="w", pady=(0, 10))
        grid_f = ttk.Frame(c2, style="Card.TFrame")
        grid_f.pack(fill="x")
        labels = [("Start Date", self.start_var), ("Benchmark", self.bench_var), ("Rebalancing", self.rebalance_var), ("Currency", self.currency_var)]
        for i, (txt, var, vals) in enumerate([(l, v, (["SPY", "QQQ", "EWY", "VTI", "069500.KS"] if l=="Benchmark" else (["Monthly", "Quarterly", "Yearly", "None"] if l=="Rebalancing" else (["KRW", "USD"] if l=="Currency" else None)))) for l, v in labels]):
            ttk.Label(grid_f, text=txt, style="CardLabel.TLabel").grid(row=i, column=0, sticky="w", pady=5)
            if vals:
                ttk.Combobox(grid_f, textvariable=var, values=vals, state="readonly", width=13).grid(row=i, column=1, sticky="e")
            else:
                ttk.Entry(grid_f, textvariable=var, width=15).grid(row=i, column=1, sticky="e")

        main_area = ttk.Frame(root)
        main_area.grid(row=0, column=1, sticky="nsew")
        main_area.rowconfigure(1, weight=1)
        manage_f = ttk.Frame(main_area, style="Card.TFrame", padding=10)
        manage_f.pack(fill="x", pady=(0, 15))
        ttk.Label(manage_f, text="Strategy Portfolio Management", style="Header.TLabel").pack(side="left", padx=5)
        ttk.Button(manage_f, text="📂 Load Portfolio", width=15, command=self._load_config).pack(side="right", padx=5)
        ttk.Button(manage_f, text="💾 Save Current Portfolio", width=20, style="Accent.TButton", command=self._save_config).pack(side="right", padx=5)

        ports_f = ttk.Frame(main_area)
        ports_f.pack(fill="both", expand=True)
        ports_f.columnconfigure(0, weight=1); ports_f.columnconfigure(1, weight=1)

        def build_port_box(parent, title, which):
            frame = ttk.Frame(parent, style="Card.TFrame", padding=15)
            ttk.Label(frame, text=title, style="Header.TLabel").pack(anchor="w", pady=(0, 10))
            tree = ttk.Treeview(frame, columns=("ticker", "name", "weight"), show="headings", height=10)
            tree.heading("ticker", text="Ticker"); tree.heading("name", text="Asset Name"); tree.heading("weight", text="W(%)")
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
            ttk.Button(btn_box, text="Del", width=5, command=lambda: self._delete_asset(which)).pack(side="left")
            ttk.Button(btn_box, text="Equal", width=8, command=lambda: self._equal_weight(which)).pack(side="left", padx=5)
            ttk.Button(btn_box, text="Reset", width=7, command=lambda: self._reset_weights(which)).pack(side="left")
            sum_lbl = ttk.Label(btn_box, text="Sum: 0%", style="CardLabel.TLabel"); sum_lbl.pack(side="right")
            if which == "A": self.tree_a, self.edit_a, self.sum_a, self.box_a = tree, edit_inner, sum_lbl, frame
            else: self.tree_b, self.edit_b, self.sum_b, self.box_b = tree, edit_inner, sum_lbl, frame
            return frame

        build_port_box(ports_f, "Portfolio A (Strategy)", "A").grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        build_port_box(ports_f, "Portfolio B (Comparative)", "B").grid(row=0, column=1, sticky="nsew")
        ttk.Button(main_area, text="🚀 RUN LONG-TERM BACKTEST", style="Success.TButton", command=self._run).pack(fill="x", pady=(15, 0), ipady=12)

    # --- LOGIC METHODS ---
    def _start_universe_drag(self, event):
        iid = self.asset_tree.identify_row(event.y)
        if iid: 
            vals = self.asset_tree.item(iid)['values']
            if vals: self.drag_data = {"source": "universe", "ticker": str(vals[0]), "port": None, "iid": iid}; self.asset_tree.config(cursor="hand2")
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
            target_dict = self.port_a if drop_target == "A" else self.port_b
            if ticker not in target_dict: target_dict[ticker] = 0.0; self._refresh_port_ui(drop_target)
        elif source == "portfolio" and drop_target == self.drag_data["port"]:
            tree = self.tree_a if drop_target == "A" else self.tree_b
            rel_y = y - tree.winfo_rooty(); drop_iid = tree.identify_row(rel_y)
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
            item = self.asset_tree.item(sel[0])
            if item['values']: 
                ticker = str(item['values'][0]); target = self.port_a if which == "A" else self.port_b
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
        
        # 상단 Treeview와 동일한 컬럼 구조 (Ticker, Name, Weight)
        edit.columnconfigure(0, minsize=70) # Ticker 컬럼
        edit.columnconfigure(1, minsize=200) # Name 컬럼
        edit.columnconfigure(2, minsize=60)  # Weight 컬럼
        
        # 1. 헤더 (생략 가능하나 가독성을 위해 유지)
        ttk.Label(edit, text="Ticker", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w", padx=5)
        ttk.Label(edit, text="Asset Name", style="CardLabel.TLabel").grid(row=0, column=1, sticky="w", padx=5)
        ttk.Label(edit, text="W(%)", style="CardLabel.TLabel").grid(row=0, column=2, sticky="e", padx=5)
        
        total_w = 0
        for i, (t, w) in enumerate(data.items()):
            full_name = backtest_proxy.get_asset_name(t)
            short_name = full_name.split(' (')[0]
            tree.insert("", "end", values=(t, short_name, f"{w*100:.1f}")); total_w += w
            
            # 1열: Ticker (중앙)
            ttk.Label(edit, text=t, style="CardLabel.TLabel").grid(row=i+1, column=0, sticky="w", padx=5)
            # 2열: Asset Name (왼쪽)
            ttk.Label(edit, text=short_name[:18], style="CardLabel.TLabel").grid(row=i+1, column=1, sticky="w", padx=5)
            # 3열: Weight 입력창 (오른쪽 정렬, 상단 Treeview 비중 컬럼과 수직 정렬)
            var = tk.StringVar(value=f"{w*100:.1f}")
            ent = ttk.Entry(edit, textvariable=var, width=8, justify="right")
            ent.grid(row=i+1, column=2, sticky="e", padx=5, pady=2)
            var.trace_add("write", lambda *args, t=t, v=var, w=which: self._update_weight(w, t, v))
        
        sum_lbl.config(text=f"Sum: {total_w*100:.1f}%")
    def _update_weight(self, which, ticker, var):
        try:
            val = float(var.get() or 0) / 100
            if which == "A": self.port_a[ticker] = val
            else: self.port_b[ticker] = val
            data = self.port_a if which == "A" else self.port_b; sum_lbl = self.sum_a if which == "A" else self.sum_b
            sum_lbl.config(text=f"Sum: {sum(data.values())*100:.1f}%")
        except: pass
    def _save_config(self):
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if path:
            try:
                with open(path, 'w', encoding='utf-8') as f:
                    json.dump({"start": self.start_var.get(), "bench": self.bench_var.get(), "reb": self.rebalance_var.get(), "cur": self.currency_var.get(), "a": self.port_a, "b": self.port_b}, f, indent=4, ensure_ascii=False)
                messagebox.showinfo("Success", "Saved.")
            except Exception as e: messagebox.showerror("Error", str(e))
    def _load_config(self):
        path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if path:
            try:
                with open(path, 'r', encoding='utf-8') as f: cfg = json.load(f)
                self.start_var.set(cfg.get("start", "2005-01-01")); self.bench_var.set(cfg.get("bench", "SPY")); self.rebalance_var.set(cfg.get("reb", "Monthly")); self.currency_var.set(cfg.get("cur", "KRW"))
                self.port_a, self.port_b = cfg.get("a", {}), cfg.get("b", {}); self._refresh_port_ui("A"); self._refresh_port_ui("B")
            except Exception as e: messagebox.showerror("Error", str(e))
    def _delete_asset(self, which):
        tree, data = (self.tree_a, self.port_a) if which == "A" else (self.tree_b, self.port_b)
        for item in tree.selection():
            ticker = str(tree.item(item)['values'][0])
            if ticker in data: del data[ticker]
        self._refresh_port_ui(which)
    def _equal_weight(self, which):
        data = self.port_a if which == "A" else self.port_b
        if data:
            eq = 1.0 / len(data)
            for t in data: data[t] = eq
            self._refresh_port_ui(which)

    def _reset_weights(self, which):
        data = self.port_a if which == "A" else self.port_b
        if data:
            for t in data: data[t] = 0.0
            self._refresh_port_ui(which)

    def _run(self):
        try:
            if not self.port_a or not self.port_b: raise ValueError("Both portfolios must have assets.")
            if abs(sum(self.port_a.values()) - 1.0) > 0.05 or abs(sum(self.port_b.values()) - 1.0) > 0.05: raise ValueError("Weights must sum to 100% (±5% allowed).")
            res = backtest_proxy.run_pro_backtest(self.port_a, self.port_b, start=self.start_var.get(), benchmark_ticker=self.bench_var.get(), rebalance=self.rebalance_var.get(), base_currency=self.currency_var.get())
            self._show_results(res)
        except Exception as e: messagebox.showerror("Error", str(e))

    def _show_results(self, res):
        win = tk.Toplevel(self); win.title("Professional Backtest Report"); win.geometry("1300x950"); win.configure(bg=COLOR_BG)
        top = ttk.Frame(win, padding=(20, 10)); top.pack(fill="x")
        curr = self.currency_var.get(); ttk.Label(top, text=f"Performance Analysis ({curr} Base)", font=FONT_TITLE).pack(side="left")
        
        def export_excel():
            path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if path:
                try:
                    with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
                        res['metrics'].to_excel(writer, sheet_name='Summary')
                        for p_idx, p_data in [('A', self.port_a), ('B', self.port_b)]:
                            valid_tickers = [t for t in p_data.keys() if t in res['raw_returns'].columns]
                            if valid_tickers:
                                detail_dfs = []
                                fx_data = res['raw_prices']['KRW=X'] if 'KRW=X' in res['raw_prices'].columns else pd.Series(1.0, index=res['raw_prices'].index)
                                for t in valid_tickers:
                                    t_df = pd.DataFrame(index=res['raw_returns'].index)
                                    t_df[f'{t}_Market_Price'] = res['raw_prices'][t] if t in res['raw_prices'].columns else np.nan
                                    t_df[f'Exchange_Rate'] = fx_data
                                    t_df[f'{t}_Value_100M'] = (1 + res['raw_returns'][t]).cumprod() * 100_000_000
                                    detail_dfs.append(t_df)
                                pd.concat(detail_dfs, axis=1).to_excel(writer, sheet_name=f'Price_Detail_{p_idx}')
                        res['monthly_a'].to_excel(writer, sheet_name='Monthly_A'); res['monthly_b'].to_excel(writer, sheet_name='Monthly_B')
                        res['corr_a'].to_excel(writer, sheet_name='Corr_A'); res['corr_b'].to_excel(writer, sheet_name='Corr_B')
                    messagebox.showinfo("Success", "Excel report saved.")
                except Exception as e: messagebox.showerror("Error", str(e))

        ttk.Button(top, text="📥 Export Excel", style="Success.TButton", command=export_excel).pack(side="right")
        nb = ttk.Notebook(win); nb.pack(fill="both", expand=True, padx=20, pady=10)

        def create_scroll_tab(notebook, title):
            frame = ttk.Frame(notebook); notebook.add(frame, text=title)
            canvas = tk.Canvas(frame, bg=COLOR_BG, highlightthickness=0)
            scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
            scrollable_inner = ttk.Frame(canvas, style="Card.TFrame")
            
            # 중앙 정렬을 위한 캔버스 설정 보강
            def _on_canvas_configure(event):
                canvas.coords(canvas_window, event.width / 2, 0)
            
            canvas.bind("<Configure>", _on_canvas_configure)
            scrollable_inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
            
            # anchor="n" (중앙 상단)으로 윈도우 생성
            canvas_window = canvas.create_window((650, 0), window=scrollable_inner, anchor="n")
            
            canvas.configure(yscrollcommand=scrollbar.set)
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            def _on_mousewheel(event):
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            
            canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
            canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
            return scrollable_inner

        # 모든 차트 크기 통일 및 확장
        UNIFIED_FIG = (15, 15)

        # Tab 1: Stats
        t1_inner = create_scroll_tab(nb, "Performance Stats")
        tree = ttk.Treeview(t1_inner, columns=("Metric", "Port A", "Port B", "Benchmark"), show="headings", height=15)
        for c in tree["columns"]: tree.heading(c, text=c); tree.column(c, anchor="center", width=250)
        tree.pack(pady=30, padx=30, anchor="n"); 
        for idx, row in res['metrics'].iterrows(): tree.insert("", "end", values=(idx, row['Port A'], row['Port B'], row['Benchmark']))

        # Tab 2: Growth & Risk
        t2_inner = create_scroll_tab(nb, "Growth & Risk")
        fig = Figure(figsize=UNIFIED_FIG); axs = fig.subplots(2, 1); bench = self.bench_var.get()
        axs[0].plot(res['asset_values_a'], label="Port A", lw=3); axs[0].plot(res['asset_values_b'], label="Port B", lw=3); axs[0].plot(res['asset_values_bench'], label=f"Bench({bench})", ls="--", alpha=0.7)
        axs[0].set_title(f"Cumulative Equity Growth ({curr})", fontsize=16, pad=30, fontweight='bold'); axs[0].legend(fontsize=12); axs[0].grid(True, ls=":", alpha=0.6)
        axs[0].set_yscale('linear'); from matplotlib.ticker import FuncFormatter, MaxNLocator
        axs[0].yaxis.set_major_formatter(FuncFormatter(lambda x, p: format(int(x), ','))); axs[0].yaxis.set_major_locator(MaxNLocator(nbins=12))
        axs[1].fill_between(res['drawdown_a'].index, res['drawdown_a'], alpha=0.15); axs[1].plot(res['drawdown_a'], label="A DD"); axs[1].plot(res['drawdown_b'], label="B DD"); axs[1].plot(res['drawdown_bench'], label="Bench DD", ls="--"); axs[1].legend(fontsize=12); axs[1].grid(True, ls=":", alpha=0.6)
        axs[1].set_title("Drawdown Analysis (%)", fontsize=16, pad=30, fontweight='bold')
        fig.tight_layout(pad=8.0); FigureCanvasTkAgg(fig, master=t2_inner).get_tk_widget().pack(fill="x", pady=20, anchor="n")

        # Tab 3: Asset Correlation (가로 확장 및 중앙 정렬 강화)
        t3_inner = create_scroll_tab(nb, "Asset Correlation")
        fig_c = Figure(figsize=UNIFIED_FIG); axs_c = fig_c.subplots(2, 1)
        def draw_corr(ax, df, title):
            if not df.empty:
                # aspect='auto'를 적용하여 가로로 길게 확장 (Monthly Returns 탭과 동일한 설정)
                im = ax.imshow(df.values, cmap='RdYlGn', vmin=-1, vmax=1, aspect='auto')
                ax.set_title(title, pad=35, fontsize=16, fontweight='bold')
                ax.set_xticks(range(len(df.columns))); ax.set_yticks(range(len(df.index)))
                ax.set_xticklabels(df.columns, rotation=45, ha='right', fontsize=11); ax.set_yticklabels(df.index, fontsize=11)
                for i in range(len(df.index)):
                    for j in range(len(df.columns)): ax.text(j, i, f"{df.iloc[i,j]:.2f}", ha="center", va="center", fontsize=11, fontweight='bold')
                fig_c.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
        draw_corr(axs_c[0], res['corr_a'], "Portfolio A Asset Correlation"); draw_corr(axs_c[1], res['corr_b'], "Portfolio B Asset Correlation")
        fig_c.tight_layout(pad=10.0); FigureCanvasTkAgg(fig_c, master=t3_inner).get_tk_widget().pack(fill="x", pady=20, anchor="n")

        # Tab 4: Monthly Returns
        t4_inner = create_scroll_tab(nb, "Monthly Returns")
        fig_m = Figure(figsize=UNIFIED_FIG); axs_m = fig_m.subplots(2, 1)
        def draw_m(ax, df, title):
            if not df.empty:
                im = ax.imshow(df.values, cmap='RdYlGn', vmin=-5, vmax=5, aspect='auto'); ax.set_title(title, pad=35, fontsize=16, fontweight='bold')
                ax.set_yticks(range(len(df.index))); ax.set_yticklabels(df.index, fontsize=12); ax.set_xticks(range(len(df.columns))); ax.set_xticklabels([f"{m}M" for m in df.columns], fontsize=12)
                for i in range(len(df.index)):
                    for j in range(len(df.columns)):
                        val = df.iloc[i, j]
                        if pd.notna(val):
                            txt = ax.annotate(f"{float(val):.1f}", xy=(j, i), ha="center", va="center", color="black", fontsize=11, fontweight='bold')
                            txt.set_path_effects([patheffects.withStroke(linewidth=2, foreground='white')])
                fig_m.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
        draw_m(axs_m[0], res['monthly_a'], "Portfolio A Monthly Returns (%)"); draw_m(axs_m[1], res['monthly_b'], "Portfolio B Monthly Returns (%)")
        fig_m.tight_layout(pad=10.0); FigureCanvasTkAgg(fig_m, master=t4_inner).get_tk_widget().pack(fill="x", pady=20, anchor="n")

if __name__ == "__main__": ProBacktestGUI().mainloop()
