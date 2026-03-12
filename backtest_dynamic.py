import numpy as np
import pandas as pd
import yfinance as yf
from datetime import datetime
import backtest_proxy

def fetch_dynamic_data(tickers, start_date, base_currency="KRW"):
    """자산 데이터를 가져와 기준 통화로 변환 (수정주가 사용)"""
    try:
        data = yf.download(tickers, start=start_date, auto_adjust=True, progress=False)['Close']
    except Exception as e:
        raise ConnectionError(f"데이터 다운로드 실패: {e}")
        
    if isinstance(data, pd.Series): data = data.to_frame()
    data = data.ffill().dropna()
    returns = data.pct_change().dropna()
    
    if base_currency == "KRW":
        usd_krw_data = yf.download("USDKRW=X", start=start_date, auto_adjust=True, progress=False)['Close']
        if isinstance(usd_krw_data, pd.Series): usd_krw = usd_krw_data.ffill()
        else: usd_krw = usd_krw_data.iloc[:, 0].ffill()
        
        usd_krw = usd_krw.reindex(returns.index, method='ffill')
        fx_ret = usd_krw.pct_change().fillna(0)
        
        for col in returns.columns:
            if not (col.endswith(".KS") or col.endswith(".KQ")):
                returns[col] = (1 + returns[col]) * (1 + fx_ret) - 1
    return returns, data

def get_weighted_momentum_score(returns_series):
    """가중 모멘텀 점수 계산 (12*1m + 4*3m + 2*6m + 1*12m)"""
    if len(returns_series) < 12: return -1
    r1 = (1 + returns_series.iloc[-1]) - 1
    r3 = (1 + returns_series.iloc[-3:]).prod() - 1
    r6 = (1 + returns_series.iloc[-6:]).prod() - 1
    r12 = (1 + returns_series.iloc[-12:]).prod() - 1
    return (12 * r1) + (4 * r3) + (2 * r6) + (1 * r12)

def run_dynamic_strategy(
    strategy_type, 
    offensive_universe, 
    defensive_universe,
    canary_universe=None, 
    start="2010-01-01", 
    initial_investment=300_000_000, 
    top_n=2, 
    base_currency="KRW",
    monthly_contribution=0,
    benchmark_ticker="SPY"
):
    """다중 동적 자산배분 엔진"""
    all_tickers = list(set(offensive_universe + defensive_universe + (canary_universe if canary_universe else []) + ["SPY", "VEA", benchmark_ticker]))
    start_dt = datetime.strptime(start, "%Y-%m-%d")
    fetch_start = (start_dt - pd.DateOffset(months=14)).strftime("%Y-%m-%d")
    
    returns, _ = fetch_dynamic_data(all_tickers, fetch_start, base_currency)
    analysis_index = returns.loc[start:].index
    if analysis_index.empty: raise ValueError("분석 기간에 해당하는 데이터가 없습니다.")
    
    current_value = float(initial_investment)
    total_invested = float(initial_investment)
    values, invested_h, weights_h = [], [], []

    for i in range(len(analysis_index)):
        curr_date = analysis_index[i]
        curr_idx = returns.index.get_loc(curr_date)
        past_returns = returns.iloc[curr_idx - 13 : curr_idx]
        target_w = pd.Series(0.0, index=all_tickers)
        
        if strategy_type == 'VAA':
            o_scores = pd.Series({t: get_weighted_momentum_score(past_returns[[t]]) for t in offensive_universe})
            if (o_scores <= 0).any():
                d_scores = pd.Series({t: get_weighted_momentum_score(past_returns[[t]]) for t in defensive_universe})
                target_w[d_scores.idxmax()] = 1.0
            else:
                top = o_scores.sort_values(ascending=False).head(top_n).index
                for t in top: target_w[t] = 1.0 / top_n
        elif strategy_type == 'DAA':
            c_scores = pd.Series({t: get_weighted_momentum_score(past_returns[[t]]) for t in canary_universe})
            num_bad = (c_scores <= 0).sum()
            off_w = 1.0 if num_bad == 0 else (0.5 if num_bad == 1 else 0.0)
            if off_w > 0:
                o_scores = pd.Series({t: get_weighted_momentum_score(past_returns[[t]]) for t in offensive_universe})
                top = o_scores.sort_values(ascending=False).head(top_n).index
                for t in top: target_w[t] = off_w / top_n
            if off_w < 1.0:
                d_scores = pd.Series({t: get_weighted_momentum_score(past_returns[[t]]) for t in defensive_universe})
                target_w[d_scores.idxmax()] += (1.0 - off_w)
        elif strategy_type == 'GEM':
            spy_ret = (1 + past_returns["SPY"].iloc[-12:]).prod() - 1
            if spy_ret > 0:
                vea_ret = (1 + past_returns["VEA"].iloc[-12:]).prod() - 1
                target_w["SPY" if spy_ret > vea_ret else "VEA"] = 1.0
            else: target_w[defensive_universe[0]] = 1.0

        monthly_ret = (returns.loc[curr_date, all_tickers] * target_w).sum()
        current_value = current_value * (1 + monthly_ret) + monthly_contribution
        total_invested += monthly_contribution
        values.append(current_value); invested_h.append(total_invested); weights_h.append(target_w)

    val_s = pd.Series(values, index=analysis_index)
    bench_ret = returns.loc[analysis_index, benchmark_ticker]
    v_bench = []; tmp_v = float(initial_investment)
    for r in bench_ret: tmp_v = tmp_v * (1 + r) + monthly_contribution; v_bench.append(tmp_v)
    
    def get_m(v, i):
        final, inv = v.iloc[-1], i.iloc[-1]
        years = len(v) / 12
        cagr = (final / inv) ** (1/years) - 1 if years > 0 else 0
        mdd = (v / v.cummax() - 1).min()
        return {"총투입": f"{inv:,.0f}", "최종가치": f"{final:,.0f}", "CAGR(%)": f"{cagr*100:.2f}", "MDD(%)": f"{mdd*100:.2f}"}

    metrics = pd.DataFrame({'Strategy': get_m(val_s, pd.Series(invested_h, index=analysis_index)), 'Benchmark': get_m(pd.Series(v_bench, index=analysis_index), pd.Series(invested_h, index=analysis_index))})
    
    # Robust Monthly Matrix
    rets = val_s.pct_change().fillna(0)
    matrix = (rets.to_frame('ret').assign(year=rets.index.year, month=rets.index.month)
              .groupby(['year', 'month'])['ret'].apply(lambda x: (1+x).prod()-1)*100).unstack()
    matrix = matrix.reindex(columns=range(1, 13))
    matrix.columns = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

    return {
        "asset_values": val_s, "asset_values_bench": pd.Series(v_bench, index=analysis_index),
        "metrics": metrics, "weights": pd.DataFrame(weights_h, index=analysis_index),
        "drawdown": (val_s / val_s.cummax() - 1) * 100, "monthly_matrix": matrix
    }

if __name__ == "__main__":
    off = ["SPY", "VEA", "VWO", "AGG"]
    dfn = ["SHY", "BIL", "IEF"]
    res = run_dynamic_strategy('VAA', off, dfn, start="2015-01-01")
    print("=== Dynamic Logic Test Result ===")
    print(res['metrics'])
