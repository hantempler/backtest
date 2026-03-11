import numpy as np
import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt
from datetime import datetime
import os

# =================================================================
# 1. 자산배분 유니버스 및 프리셋 (Professional Universe & Presets)
# =================================================================
ASSET_UNIVERSE = {
    "주식 (Stocks)": {
        "SPY": "미국 S&P 500 (SPDR S&P 500)",
        "VTI": "미국 주식 전체 (Vanguard Total Stock)",
        "QQQ": "미국 나스닥 100 (Invesco QQQ)",
        "EWY": "한국 주식 (iShares MSCI South Korea)",
        "VEA": "선진국 주식 (Vanguard FTSE Dev.)",
        "VWO": "신흥국 주식 (Vanguard FTSE EM)",
        "SCHD": "미국 배당성장 (Schwab US Dividend)",
        "ASHR": "중국 CSI 300 (Xtrackers China)",
        "INDA": "인도 Nifty 50 (iShares India)"
    },
    "채권 (Bonds)": {
        "TLT": "미국 장기 국채 (20Y+ Treasury)",
        "IEF": "미국 중기 국채 (7-10Y Treasury)",
        "BND": "미국 종합 채권 (Vanguard Total Bond)",
        "152380.KS": "한국 국고채 10년 (KODEX)",
        "TIP": "미국 물가 연동채 (iShares TIPS)",
        "LQD": "미국 투자등급 회사채 (iShares Corp Bond)"
    },
    "대체자산 (Alternatives)": {
        "GLD": "금 (SPDR Gold Shares)",
        "DBC": "원자재 종합 (Invesco DB Commodity)",
        "VNQ": "미국 리츠 (Vanguard Real Estate)",
        "RWX": "글로벌 리츠 (SPDR Intl Real Estate)"
    },
    "현금/단기채 (Cash)": {
        "SHY": "미국 단기 국채 (1-3Y Treasury)",
        "BIL": "미국 초단기 국채 (1-3M Treasury)",
        "153130.KS": "한국 단기 채권 (KODEX)"
    }
}

STRATEGY_PRESETS = {
    "글로벌 60/40": {
        "VTI": 0.60,
        "BND": 0.40
    },
    "오리지널 올웨더 (Dalio)": {
        "VTI": 0.30,
        "TLT": 0.40,
        "IEF": 0.15,
        "GLD": 0.075,
        "DBC": 0.075
    },
    "영구 포트폴리오 (Browne)": {
        "VTI": 0.25,
        "TLT": 0.25,
        "GLD": 0.25,
        "BIL": 0.25
    },
    "골든 버터플라이": {
        "VTI": 0.20,
        "SCHD": 0.20,
        "TLT": 0.20,
        "SHY": 0.20,
        "GLD": 0.20
    },
    "한국형 글로벌 자산배분": {
        "EWY": 0.20,
        "VTI": 0.20,
        "152380.KS": 0.30,
        "TLT": 0.20,
        "153130.KS": 0.10
    }
}

def is_us_proxy(ticker):
    t = str(ticker).upper()
    return not t.endswith(".KS") and t != "KRW=X" and not t.isdigit()

def get_asset_name(ticker):
    for cat, assets in ASSET_UNIVERSE.items():
        if ticker in assets: return assets[ticker]
    return f"Custom({ticker})"

def fetch_hybrid_data(tickers, start_date, base_currency="KRW"):
    download_list = list(set(tickers + ["KRW=X"]))
    try:
        raw_data = yf.download(download_list, start=start_date, progress=False)['Close']
        if raw_data.empty:
            raise ValueError("가져온 가격 데이터가 비어 있습니다.")
    except Exception as e:
        raise ConnectionError(f"Yahoo Finance 데이터 다운로드 실패: {str(e)}")

    monthly = raw_data.resample('ME').last()
    returns = monthly.pct_change().dropna(how='all')
    fx_ret = returns['KRW=X'] if 'KRW=X' in returns.columns else None
    
    final_returns = pd.DataFrame(index=returns.index)
    for t in tickers:
        if t in returns.columns:
            is_us = is_us_proxy(t)
            if base_currency == "KRW":
                if is_us and fx_ret is not None:
                    final_returns[t] = (1 + returns[t]) * (1 + fx_ret) - 1
                else:
                    final_returns[t] = returns[t]
            else: # USD Base
                if not is_us and fx_ret is not None:
                    final_returns[t] = (1 + returns[t]) / (1 + fx_ret) - 1
                else:
                    final_returns[t] = returns[t]
    
    return final_returns.dropna(axis=1, how='all').dropna(), monthly

def get_performance_metrics(returns, initial_investment=300_000_000, rf=0.0):
    if returns.empty: return {}
    cum = (1 + returns).cumprod()
    total_ret = cum.iloc[-1] - 1
    years = len(returns) / 12
    cagr = (cum.iloc[-1] ** (1/years)) - 1 if years > 0 else 0
    vol = returns.std() * np.sqrt(12)
    sharpe = (cagr - rf) / vol if vol > 0 else 0
    mdd = (cum / cum.cummax() - 1).min()
    return {
        "누적 수익률(%)": f"{total_ret*100:.2f}",
        "CAGR(%)": f"{cagr*100:.2f}",
        "연변동성(%)": f"{vol*100:.2f}",
        "샤프비율": f"{sharpe:.2f}",
        "MDD(%)": f"{mdd*100:.2f}"
    }

def get_monthly_matrix(ret_series):
    df = ret_series.to_frame(name='ret')
    df['year'], df['month'] = df.index.year, df.index.month
    return df.pivot_table(index='year', columns='month', values='ret') * 100

def run_pro_backtest(port_a_w, port_b_w, start="2010-01-01", initial_investment=300_000_000, benchmark_ticker="SPY", rebalance="Monthly", base_currency="KRW"):
    all_tickers = list(set(list(port_a_w.keys()) + list(port_b_w.keys()) + [benchmark_ticker]))
    returns, raw_monthly_prices = fetch_hybrid_data(all_tickers, start, base_currency)
    
    avail_tickers = returns.columns.tolist()
    
    def calc_port_ret(w_dict, rebalance_type):
        valid_w = {t: v for t, v in w_dict.items() if t in avail_tickers}
        if not valid_w: return pd.Series(0.0, index=returns.index)
        w = pd.Series(0.0, index=avail_tickers)
        for t, val in valid_w.items(): w[t] = val
        if w.sum() > 0: w /= w.sum()
        
        if rebalance_type == "Monthly":
            return (returns[avail_tickers] * w).sum(axis=1)
        elif rebalance_type == "None":
            asset_values = (1 + returns[avail_tickers]).cumprod().mul(w, axis=1)
            port_values = asset_values.sum(axis=1)
            return port_values.pct_change().fillna(0)
        else:
            step = 3 if rebalance_type == "Quarterly" else 12
            port_returns = []
            curr_w = w.values
            for i in range(len(returns)):
                if i > 0 and i % step == 0: curr_w = w.values
                ret_row = returns.iloc[i].values
                p_ret = np.dot(ret_row, curr_w)
                port_returns.append(p_ret)
                curr_w = curr_w * (1 + ret_row)
                if curr_w.sum() > 0: curr_w /= curr_w.sum()
            return pd.Series(port_returns, index=returns.index)

    ret_a = calc_port_ret(port_a_w, rebalance)
    ret_b = calc_port_ret(port_b_w, rebalance)
    ret_bench = returns[benchmark_ticker] if benchmark_ticker in returns.columns else pd.Series(0.0, index=returns.index)

    val_a = (1 + ret_a).cumprod() * initial_investment
    val_b = (1 + ret_b).cumprod() * initial_investment
    val_bench = (1 + ret_bench).cumprod() * initial_investment

    metrics = pd.DataFrame({
        'Port A': get_performance_metrics(ret_a, initial_investment),
        'Port B': get_performance_metrics(ret_b, initial_investment),
        'Benchmark': get_performance_metrics(ret_bench, initial_investment)
    })

    def get_corr(w_dict):
        tickers = [t for t, v in w_dict.items() if v > 0 and t in avail_tickers]
        if len(tickers) < 2: return pd.DataFrame()
        corr = returns[tickers].corr()
        corr.index = [get_asset_name(t).split(' (')[0] for t in corr.index]
        corr.columns = [get_asset_name(t).split(' (')[0] for t in corr.columns]
        return corr

    return {
        "asset_values_a": val_a, "asset_values_b": val_b, "asset_values_bench": val_bench,
        "metrics": metrics,
        "monthly_a": get_monthly_matrix(ret_a), "monthly_b": get_monthly_matrix(ret_b),
        "corr_a": get_corr(port_a_w), "corr_b": get_corr(port_b_w),
        "drawdown_a": (val_a / val_a.cummax() - 1) * 100,
        "drawdown_b": (val_b / val_b.cummax() - 1) * 100,
        "drawdown_bench": (val_bench / val_bench.cummax() - 1) * 100,
        "raw_returns": returns,
        "raw_prices": raw_monthly_prices # 정확한 변수명으로 매핑
    }
