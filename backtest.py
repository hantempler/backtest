import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import yfinance as yf
from matplotlib.ticker import FuncFormatter
import os

plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

# ==============================
# 1. ETF 정보 로드 및 컬럼 정규화
# ==============================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
etf_data = os.path.join(BASE_DIR, "etf_data.xlsx")

df_excel = pd.read_excel(etf_data)

# 컬럼 이름 정규화 (실제 엑셀 열 이름에 맞게 매핑)
df_master = df_excel.copy()
df_master = df_master.rename(columns={
    '단축코드': 'ticker_code',
    '한글종목약명': 'name',
    '기초자산분류': 'group',
    '기초시장분류': 'sub',           # ← 실제 열 이름이 '기초시장분류'라면 이렇게
    '상장일': 'listed'
})

# ticker 정제
df_master['ticker'] = df_master['ticker_code'].astype(str).str.strip().str.zfill(6)
df_master['ticker_ks'] = df_master['ticker'] + ".KS"

# 중복/불필요 컬럼 정리
df_master = df_master.drop(columns=['ticker_code'], errors='ignore')

# 필수 컬럼 확인
required = ['ticker', 'ticker_ks', 'name', 'group', 'sub']
missing = [c for c in required if c not in df_master.columns]
if missing:
    raise ValueError(f"엑셀 파일에 필수 컬럼이 없습니다: {missing}")


def make_equal_weight(df_m, selected_tickers=None):
    df = df_m.copy()
    if selected_tickers is not None:
        df = df[df['ticker'].isin(map(str, selected_tickers))]
    if len(df) == 0:
        raise ValueError("선택된 티커가 없습니다.")
    w = 1.0 / len(df)
    return {t: w for t in df['ticker'].tolist()}


def make_group_weighted(df_m, weights_by_group, selected_tickers=None):
    df = df_m.copy()
    if selected_tickers is not None:
        df = df[df['ticker'].isin(map(str, selected_tickers))]
    if len(df) == 0:
        raise ValueError("선택된 티커가 없습니다.")

    groups_present = df['group'].dropna().unique().tolist()

    # 없는 그룹은 자동으로 0 비중 처리
    active_weights = {}
    total_weight = 0.0
    for g in groups_present:
        w = float(weights_by_group.get(g, 0.0))
        active_weights[g] = w
        total_weight += w

    if total_weight <= 0:
        raise ValueError("유효한 그룹 비중의 합이 0입니다.")

    # 정규화 (0인 그룹은 제외)
    norm_weights = {g: w / total_weight for g, w in active_weights.items() if w > 0}

    group_counts = df.groupby('group').size().to_dict()
    out = {}
    for _, row in df.iterrows():
        g = row['group']
        if g in norm_weights and group_counts.get(g, 0) > 0:
            out[str(row['ticker'])] = norm_weights[g] / group_counts[g]
        else:
            out[str(row['ticker'])] = 0.0

    return out


def build_df_weights(df_m, port_a, port_b):
    port_a = {str(k): float(v) for k, v in port_a.items()}
    port_b = {str(k): float(v) for k, v in port_b.items()}

    selected = sorted(set(port_a) | set(port_b))
    df = df_m[df_m['ticker'].isin(selected)].copy()

    missing = [t for t in selected if t not in df['ticker'].values]
    if missing:
        raise ValueError(f"asset_master에 없는 티커: {missing}")

    df['port_A'] = df['ticker'].map(port_a).fillna(0.0)
    df['port_B'] = df['ticker'].map(port_b).fillna(0.0)

    sum_a = df['port_A'].sum()
    sum_b = df['port_B'].sum()
    if sum_a <= 0:
        raise ValueError("port_A 비중 합이 0 이하입니다.")
    if sum_b <= 0:
        raise ValueError("port_B 비중 합이 0 이하입니다.")

    df['port_A'] /= sum_a
    df['port_B'] /= sum_b

    out_cols = ['ticker', 'ticker_ks', 'name', 'port_A', 'port_B']
    if 'listed' in df.columns:
        out_cols.insert(out_cols.index('name') + 1, 'listed')

    return df[out_cols].sort_values('ticker')


def pick_selected_tickers(df_m, tickers=None, groups=None, subs=None):
    if tickers is None and groups is None and subs is None:
        return None

    df = df_m.copy()
    if tickers is not None:
        df = df[df['ticker'].isin(map(str, tickers))]
    if groups is not None:
        df = df[df['group'].isin(groups)]
    if subs is not None:
        df = df[df['sub'].isin(subs)]

    selected = df['ticker'].drop_duplicates().tolist()
    if not selected:
        raise ValueError("선택 조건에 맞는 티커가 없습니다.")
    return selected


# ==============================
# 사용자 설정 영역 (여기만 수정)
# ==============================

# 전체 중에서 원하는 그룹만 사용하려면 아래처럼 제한
SELECTED_GROUPS_A = ['주식', '채권', '현금', '대체자산']   # ← 에러 원인 해결 핵심
SELECTED_SUBS_A   = None
SELECTED_TICKERS_A = None

SELECTED_GROUPS_B = None
SELECTED_SUBS_B   = None
SELECTED_TICKERS_B = None

selected_tickers_a = pick_selected_tickers(
    df_master,
    tickers=SELECTED_TICKERS_A,
    groups=SELECTED_GROUPS_A,
    subs=SELECTED_SUBS_A,
)

selected_tickers_b = pick_selected_tickers(
    df_master,
    tickers=SELECTED_TICKERS_B,
    groups=SELECTED_GROUPS_B,
    subs=SELECTED_SUBS_B,
)


def plot_contribution_with_names(contrib_df, title):
    active = contrib_df.loc[:, contrib_df.any()]
    if active.empty:
        print(f"[{title}] 표시할 데이터 없음")
        return

    cum = active.cumsum()

    plt.figure(figsize=(13,7))
    for col in cum:
        plt.plot(cum.index, cum[col], label=col, lw=2)

    total = cum.sum(axis=1)
    plt.plot(total.index, total, 'k--', lw=3, label='전체')

    plt.title(title, fontsize=16)
    plt.ylabel('누적 기여도 (%)')
    plt.axhline(0, color='red', ls='--', alpha=0.6)
    plt.legend(bbox_to_anchor=(1.02,1), loc='upper left', title="ETF")
    plt.grid(True, ls=':')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()


# ==============================
# 2. 전략 프리셋 정의
# ==============================

STRATEGY_PRESETS = {
    "60/40 전략": {
        "069500": 0.60,  # KODEX 200
        "114470": 0.40   # KODEX 국고채3년
    },
    "올웨더(간소화)": {
        "069500": 0.30,  # 주식 (코스피200)
        "152380": 0.40,  # 장기채 (KODEX 10년국채선물)
        "114470": 0.15,  # 중기채 (KODEX 국고채3년)
        "132030": 0.075, # 금 (KODEX 골드선물)
        "138920": 0.075  # 원자재 (KODEX 구리선물 - 대용)
    },
    "영구 포트폴리오": {
        "069500": 0.25,  # 주식
        "152380": 0.25,  # 채권
        "132030": 0.25,  # 금
        "272580": 0.25   # 현금 (TIGER 단기통안채)
    }
}

def get_performance_metrics(returns, initial_investment=300_000_000, rf=0.0):
    returns = returns.dropna()
    if returns.empty:
        return {"오류": "수익률 데이터 없음"}

    cum = (1 + returns).cumprod()
    total_ret = cum.iloc[-1] - 1

    years = len(returns) / 12
    cagr = (cum.iloc[-1] ** (1/years)) - 1 if years > 0 else 0

    vol = returns.std() * np.sqrt(12)
    sharpe = (cagr - rf) / vol if vol > 0 else 0

    peak = cum.cummax()
    dd = (cum / peak - 1)
    mdd = dd.min()

    final = cum.iloc[-1] * initial_investment

    return {
        "누적 수익률(%)": f"{total_ret*100:.2f}",
        "CAGR(%)": f"{cagr*100:.2f}",
        "연변동성(%)": f"{vol*100:.2f}",
        "샤프비율": f"{sharpe:.2f}",
        "MDD(%)": f"{mdd*100:.2f}",
        "최종자산(원)": f"{final:,.0f}"
    }

def run_backtest(
    df_weights_input,
    start="2018-01-01",
    initial_investment=300_000_000,
    rf=0.0,
    rebalance="Monthly",  # "Monthly", "Quarterly", "Yearly", "None"
    benchmark_ticker="069500.KS", # 기본 KODEX 200
    show_plots=True,
    verbose=True,
):
    dfw = df_weights_input.copy()
    if 'ticker_ks' not in dfw.columns:
        dfw['ticker_ks'] = dfw['ticker'] + ".KS"

    tickers = list(set(dfw['ticker_ks'].tolist() + [benchmark_ticker]))
    name_map = dict(zip(dfw['ticker_ks'], dfw['name']))
    name_map[benchmark_ticker] = "벤치마크(KODEX 200)"

    if verbose:
        print(f"가격 데이터 다운로드 중... ({len(tickers)} 종목)")

    prices = yf.download(tickers, start=start, progress=False)['Close']
    monthly = prices.resample('ME').last()
    returns = monthly.pct_change().dropna()

    def calculate_portfolio_return(weights_series, rebalance_type):
        w = weights_series.reindex(returns.columns).fillna(0)
        
        if rebalance_type == "Monthly":
            # 매월 비중 재설정
            return returns.mul(w, axis=1).sum(axis=1)
        
        elif rebalance_type == "None":
            # 리밸런싱 없음 (Buy & Hold)
            asset_values = (1 + returns).cumprod().mul(w, axis=1)
            port_values = asset_values.sum(axis=1)
            return port_values.pct_change().fillna(0)
        
        else:
            # Quarterly or Yearly 리밸런싱
            step = 3 if rebalance_type == "Quarterly" else 12
            port_returns = []
            current_w = w.values
            
            for i in range(len(returns)):
                # 리밸런싱 시점 체크
                if i > 0 and i % step == 0:
                    current_w = w.values
                
                ret_row = returns.iloc[i].values
                period_ret = np.dot(ret_row, current_w)
                port_returns.append(period_ret)
                
                # 비중 변동 반영 (가격 변화에 따른 자연적 비중 변화)
                current_w = current_w * (1 + ret_row)
                current_w = current_w / current_w.sum()
                
            return pd.Series(port_returns, index=returns.index)

    w_a = dfw.set_index('ticker_ks')['port_A']
    w_b = dfw.set_index('ticker_ks')['port_B']

    port_a_ret = calculate_portfolio_return(w_a, rebalance)
    port_b_ret = calculate_portfolio_return(w_b, rebalance)
    bench_ret = returns[benchmark_ticker]

    value_a = (1 + port_a_ret).cumprod() * initial_investment
    value_b = (1 + port_b_ret).cumprod() * initial_investment
    value_bench = (1 + bench_ret).cumprod() * initial_investment

    metrics_a = get_performance_metrics(port_a_ret, initial_investment, rf)
    metrics_b = get_performance_metrics(port_b_ret, initial_investment, rf)
    metrics_bench = get_performance_metrics(bench_ret, initial_investment, rf)

    compare = pd.DataFrame([metrics_a, metrics_b, metrics_bench], index=['Port A', 'Port B', 'Benchmark']).T

    # 상관계수 계산
    corr_ab = port_a_ret.corr(port_b_ret)

    # --- 추가 분석 데이터 (2번 과제) ---
    
    # 1. 월별 수익률 히트맵용 데이터 (Port A 기준)
    def get_monthly_matrix(returns_series):
        df = returns_series.to_frame(name='ret')
        df['year'] = df.index.year
        df['month'] = df.index.month
        matrix = df.pivot_table(index='year', columns='month', values='ret')
        return matrix * 100  # % 단위

    # 2. 롤링 수익률 (12개월)
    def get_rolling_ret(returns_series, window=12):
        return (1 + returns_series).rolling(window).apply(np.prod, raw=True) - 1

    return {
        'df_weights': dfw,
        'asset_values_a': value_a,
        'asset_values_b': value_b,
        'asset_values_bench': value_bench,
        'metrics_compare': compare,
        'port_a_ret': port_a_ret,
        'port_b_ret': port_b_ret,
        'bench_ret': bench_ret,
        'drawdown_a': (value_a / value_a.cummax() - 1) * 100,
        'drawdown_b': (value_b / value_b.cummax() - 1) * 100,
        'drawdown_bench': (value_bench / value_bench.cummax() - 1) * 100,
        # 신규 분석 데이터
        'monthly_matrix_a': get_monthly_matrix(port_a_ret),
        'monthly_matrix_b': get_monthly_matrix(port_b_ret),
        'rolling_12m_a': get_rolling_ret(port_a_ret) * 100,
        'rolling_12m_b': get_rolling_ret(port_b_ret) * 100,
        'correlation_ab': corr_ab
    }



def cli_main():
    df_weights = get_default_df_weights()
    run_backtest(
        df_weights,
        start="2018-01-01",
        initial_investment=300_000_000,
        rf=0.0,
        show_plots=True,
        verbose=True
    )


if __name__ == "__main__":
    cli_main()