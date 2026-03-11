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

# 컬럼 이름 정규화
df_master = df_excel.copy()
df_master = df_master.rename(columns={
    '단축코드': 'ticker_code',
    '한글종목약명': 'name',
    '기초자산분류': 'group',
    '기초시장분류': 'sub',
    '상장일': 'listed'
})

# ticker 정제 (한국 종목은 6자리 숫자 유지, 그 외는 그대로 사용)
def clean_ticker(x):
    s = str(x).strip()
    return s.zfill(6) if s.isdigit() else s

df_master['ticker'] = df_master['ticker_code'].apply(clean_ticker)
df_master['ticker_ks'] = df_master.apply(lambda r: r['ticker'] + ".KS" if str(r['ticker']).isdigit() else r['ticker'], axis=1)

# 중복/불필요 컬럼 정리
df_master = df_master.drop(columns=['ticker_code'], errors='ignore')

def build_df_weights(df_m, port_a, port_b):
    """
    포트폴리오 비중 DataFrame 구축. 
    마스터 데이터에 없는 미국 티커(SPY 등)도 처리 가능하도록 개선.
    """
    port_a = {str(k): float(v) for k, v in port_a.items()}
    port_b = {str(k): float(v) for k, v in port_b.items()}
    selected = sorted(set(port_a) | set(port_b))
    
    # 마스터 데이터에서 추출
    df = df_m[df_m['ticker'].isin(selected)].copy()
    
    # 마스터에 없는 티커(미국 ETF 등)를 위한 가상 행 추가
    missing = [t for t in selected if t not in df['ticker'].values]
    for m_ticker in missing:
        new_row = {
            'ticker': m_ticker,
            'ticker_ks': m_ticker if ".KS" in m_ticker else m_ticker, # 이미 붙어있거나 미국 티커
            'name': f"외부자산({m_ticker})",
            'group': '기타',
            'sub': '해외'
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

    df['port_A'] = df['ticker'].map(port_a).fillna(0.0)
    df['port_B'] = df['ticker'].map(port_b).fillna(0.0)

    # 비중 정규화
    sum_a, sum_b = df['port_A'].sum(), df['port_B'].sum()
    if sum_a > 0: df['port_A'] /= sum_a
    if sum_b > 0: df['port_B'] /= sum_b

    return df.sort_values('ticker')

# ==============================
# 2. 오리지널 미국 ETF 전략 프리셋
# ==============================

STRATEGY_PRESETS = {
    "미국 60/40 전략": {
        "SPY": 0.60,  # S&P 500
        "AGG": 0.40   # US Aggregate Bond
    },
    "오리지널 올웨더": {
        "VTI": 0.30,  # 전주식시장
        "TLT": 0.40,  # 20년+ 장기국채
        "IEF": 0.15,  # 7-10년 중기국채
        "GLD": 0.075, # 금
        "DBC": 0.075  # 원자재
    },
    "오리지널 영구포트": {
        "VTI": 0.25,  # 주식
        "TLT": 0.25,  # 장기채
        "GLD": 0.25,  # 금
        "SHY": 0.25   # 단기채(현금대용)
    },
    "나스닥/배당성장": {
        "QQQ": 0.50,  # 나스닥 100
        "SCHD": 0.50  # 배당성장
    }
}

def get_performance_metrics(returns, initial_investment=300_000_000, rf=0.0):
    returns = returns.dropna()
    if returns.empty: return {"오류": "데이터 없음"}
    
    cum = (1 + returns).cumprod()
    total_ret = cum.iloc[-1] - 1
    years = len(returns) / 12
    cagr = (cum.iloc[-1] ** (1/years)) - 1 if years > 0 else 0
    vol = returns.std() * np.sqrt(12)
    sharpe = (cagr - rf) / vol if vol > 0 else 0
    mdd = (cum / cum.cummax() - 1).min()
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
    rebalance="Monthly",
    benchmark_ticker="069500.KS",
    show_plots=False,
    verbose=True,
):
    dfw = df_weights_input.copy()
    
    # 다운로드할 모든 티커 수집
    all_tickers = list(set(dfw['ticker_ks'].tolist() + [benchmark_ticker]))
    
    # 미국 티커(달러 자산) 판별: 숫자가 아니거나 .KS가 없는 경우
    def is_us_asset(t):
        return not (str(t).replace(".KS","").isdigit()) and t != "KRW=X"

    us_assets = [t for t in all_tickers if is_us_asset(t)]
    
    if us_assets:
        all_tickers.append("KRW=X")
        if verbose: print(f"미국 자산 감지됨 ({', '.join(us_assets)}). 환율 데이터를 포함합니다.")

    if verbose: print(f"가격 데이터 다운로드 중... ({len(all_tickers)} 종목)")
    prices = yf.download(all_tickers, start=start, progress=False)['Close']
    monthly = prices.resample('ME').last()
    returns = monthly.pct_change().dropna()

    # 미국 자산들에 대해 환율 변동 반영 (원화 기준 환노출 수익률로 변환)
    if "KRW=X" in returns.columns:
        fx_ret = returns["KRW=X"]
        for t in us_assets:
            if t in returns.columns:
                # 원화 수익률 = (1 + 달러 수익률) * (1 + 환율 변동률) - 1
                returns[t] = (1 + returns[t]) * (1 + fx_ret) - 1
        if verbose: print("모든 미국 자산의 수익률을 원화 기준(환노출)으로 변환 완료.")

    def calculate_portfolio_return(weights_series, rebalance_type):
        # weights_series의 인덱스를 returns.columns와 맞춤
        w = weights_series.copy()
        w.index = w.index.map(str)
        
        # returns의 컬럼명과 매칭을 위한 처리
        avail_cols = returns.columns
        w_final = pd.Series(0.0, index=avail_cols)
        
        for t, weight in w.items():
            if t in avail_cols: w_final[t] = weight
            elif (t + ".KS") in avail_cols: w_final[t + ".KS"] = weight

        if rebalance_type == "Monthly":
            return returns.mul(w_final, axis=1).sum(axis=1)
        
        elif rebalance_type == "None":
            asset_values = (1 + returns).cumprod().mul(w_final, axis=1)
            port_values = asset_values.sum(axis=1)
            return port_values.pct_change().fillna(0)
        
        else:
            step = 3 if rebalance_type == "Quarterly" else 12
            port_returns = []
            curr_w = w_final.values
            for i in range(len(returns)):
                if i > 0 and i % step == 0: curr_w = w_final.values
                ret_row = returns.iloc[i].values
                p_ret = np.dot(ret_row, curr_w)
                port_returns.append(p_ret)
                curr_w = curr_w * (1 + ret_row)
                if curr_w.sum() > 0: curr_w /= curr_w.sum()
            return pd.Series(port_returns, index=returns.index)

    # 포트폴리오 수익률 계산
    w_a = dfw.set_index('ticker_ks')['port_A']
    w_b = dfw.set_index('ticker_ks')['port_B']

    port_a_ret = calculate_portfolio_return(w_a, rebalance)
    port_b_ret = calculate_portfolio_return(w_b, rebalance)
    
    # 벤치마크 수익률 (위에서 이미 환율 변동이 반영됨)
    bench_ret = returns[benchmark_ticker]

    # 결과 취합
    value_a = (1 + port_a_ret).cumprod() * initial_investment
    value_b = (1 + port_b_ret).cumprod() * initial_investment
    value_bench = (1 + bench_ret).cumprod() * initial_investment

    metrics_compare = pd.DataFrame({
        'Port A': get_performance_metrics(port_a_ret, initial_investment, rf),
        'Port B': get_performance_metrics(port_b_ret, initial_investment, rf),
        'Benchmark': get_performance_metrics(bench_ret, initial_investment, rf)
    })

    def get_monthly_matrix(ret_series):
        df = ret_series.to_frame(name='ret')
        df['year'], df['month'] = df.index.year, df.index.month
        return df.pivot_table(index='year', columns='month', values='ret') * 100

    def get_rolling_ret(ret_series, window=12):
        return (1 + ret_series).rolling(window).apply(np.prod, raw=True) - 1

    # 포트폴리오 내 자산 간 상관관계 계산 (추가)
    def get_asset_corr(weights_series):
        active_tickers = weights_series[weights_series > 0].index.tolist()
        
        # 티커-이름 매핑 생성
        ticker_to_name = {}
        for t in active_tickers:
            row = dfw[dfw['ticker_ks'] == t]
            if not row.empty:
                ticker_to_name[t] = row.iloc[0]['name']
            else:
                ticker_to_name[t] = t

        # returns 컬럼명과 매칭 (ticker_ks 형태)
        cols = [t for t in active_tickers if t in returns.columns]
        
        if len(cols) > 1:
            corr = returns[cols].corr()
            # 티커를 종목명으로 변경
            corr.index = [ticker_to_name.get(t, t) for t in corr.index]
            corr.columns = [ticker_to_name.get(t, t) for t in corr.columns]
            return corr
        return pd.DataFrame()

    asset_corr_a = get_asset_corr(w_a)
    asset_corr_b = get_asset_corr(w_b)

    return {
        'df_weights': dfw,
        'asset_values_a': value_a,
        'asset_values_b': value_b,
        'asset_values_bench': value_bench,
        'metrics_compare': metrics_compare,
        'drawdown_a': (value_a / value_a.cummax() - 1) * 100,
        'drawdown_b': (value_b / value_b.cummax() - 1) * 100,
        'monthly_matrix_a': get_monthly_matrix(port_a_ret),
        'monthly_matrix_b': get_monthly_matrix(port_b_ret),
        'rolling_12m_a': get_rolling_ret(port_a_ret) * 100,
        'rolling_12m_b': get_rolling_ret(port_b_ret) * 100,
        'asset_corr_a': asset_corr_a,
        'asset_corr_b': asset_corr_b,
        'correlation_ab': port_a_ret.corr(port_b_ret)
    }

if __name__ == "__main__":
    print("백테스트 엔진 단독 실행 테스트")
