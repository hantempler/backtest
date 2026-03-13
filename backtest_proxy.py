import numpy as np
import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt
from datetime import datetime
import os

# =================================================================
# 1. 자산배분 유니버스 및 프리셋 (Professional Universe & Presets)
# =================================================================
GLOBAL_TICKER_NAME_MAP = {} # Dynamic map for Excel-loaded assets

ASSET_UNIVERSE = {
    "주식 (Stocks)": {
        "SPY": "미국 S&P 500 (SPDR S&P 500)",
        "VTI": "미국 주식 전체 (Vanguard Total Stock)",
        "QQQ": "미국 나스닥 100 (Invesco QQQ)",
        "VT": "전세계 주식 (Vanguard Total World)",
        "EWY": "한국 주식 (iShares MSCI South Korea)",
        "VEA": "선진국 주식 (Vanguard FTSE Dev.)",
        "VWO": "신흥국 주식 (Vanguard FTSE EM)",
        "SCHD": "미국 배당성장 (Schwab US Dividend)",
        "ASHR": "중국 CSI 300 (Xtrackers China)",
        "INDA": "인도 Nifty 50 (iShares India)",
        "SMH": "반도체 (VanEck Semiconductor)",
        "XLV": "헬스케어 (Health Care Sector)",
        "VYM": "미국 고배당주 (Vanguard High Dividend Yield)",
        "VIG": "미국 배당성장 우량주 (Vanguard Div Appreciation)",
        "IDV": "선진국 고배당 (iShares International Select Div)",
    },
    "채권 (Bonds)": {
        "TLT": "미국 장기 국채 (20Y+ Treasury)",
        "EDV": "미국 초장기 국채 (Vanguard Extended Duration)",
        "IEF": "미국 중기 국채 (7-10Y Treasury)",
        "BND": "미국 종합 채권 (Vanguard Total Bond)",
        "152380.KS": "한국 국고채 10년 (KODEX)",
        "TIP": "미국 물가 연동채 (iShares TIPS)",
        "LQD": "미국 투자등급 회사채 (iShares Corp Bond)",
        "JNK": "하이일드 채권 (SPDR High Yield)"
    },
    "대체자산 (Alternatives)": {
        "GLD": "금 (SPDR Gold Shares)",
        "IBIT": "비트코인 현물 (iShares Bitcoin)",
        "DBC": "원자재 종합 (Invesco DB Commodity)",
        "VNQ": "미국 리츠 (Vanguard Real Estate)",
        "RWX": "글로벌 리츠 (SPDR Intl Real Estate)",
        "SLV": "은 (iShares Silver Trust)"
    },
    "인컴/전략 (Income Strategy)": {
        "JEPI": "미국 배당 프리미엄 (JPMorgan Equity Premium)",
        "DIVO": "우량주 커버드콜 (Amplify Dividend Income)",
        "JEPI": "JPMorgan Equity Premium Income (월배당/저변동)",
        "JEPQ": "JPMorgan Nasdaq Equity Premium (나스닥 기반 월배당)",
        "DIVO": "Amplify CWP Enhanced Dividend Income (우량주+옵션)",
        "XYLD": "S&P 500 Covered Call (글로벌X 커버드콜)",
        "MAIN": "Main Street Capital (미국 대표 BDC/고배당 비즈니스)"
    },
    "현금/단기채 (Cash)": {
        "SHY": "미국 단기 국채 (1-3Y Treasury)",
        "BIL": "미국 초단기 국채 (1-3M Treasury)",
        "SHV": "미국 초단기 국채 (Short Treasury)",
        "153130.KS": "한국 단기 채권 (KODEX)"
    },
    
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
    "아이비 포트폴리오 (Faber)": {
        "VTI": 0.20,
        "VEA": 0.20,
        "BND": 0.20,
        "GLD": 0.20,
        "VNQ": 0.20
    },
    "예일 엔다우먼트 (Swensen)": {
        "VTI": 0.30,
        "VEA": 0.15,
        "VWO": 0.05,
        "VNQ": 0.20,
        "TIP": 0.15,
        "TLT": 0.15
    },
    "나스닥/배당성장 바벨": {
        "QQQ": 0.50,
        "SCHD": 0.50
    },
    "전지구적 주식 100%": {
        "VTI": 0.50,
        "VEA": 0.30,
        "VWO": 0.10,
        "EWY": 0.10
    },
    "인플레이션 헷지 전략": {
        "TIP": 0.30,
        "GLD": 0.30,
        "DBC": 0.30,
        "VNQ": 0.10
    },
    "신흥국 거인들 (KR/CN/IN)": {
        "EWY": 0.40,
        "ASHR": 0.30,
        "INDA": 0.30
    },
    "한국형 글로벌 자산배분": {
        "EWY": 0.20,
        "VTI": 0.20,
        "152380.KS": 0.30,
        "TLT": 0.20,
        "153130.KS": 0.10
    },
    "미국 배당성장 삼각편대": {
        "QQQ": 0.30,
        "SCHD": 0.40,
        "JEPI": 0.30
    },
    "배당형 올웨더 (Income Focus)": {
        "SCHD": 0.30,
        "TLT": 0.40,
        "LQD": 0.15,
        "GLD": 0.075,
        "JEPI": 0.075
    },
    "글로벌 고배당 인컴": {
        "VYM": 0.30,
        "IDV": 0.20,
        "VNQ": 0.20,
        "JEPI": 0.20,
        "MAIN": 0.10
    },
    "보수적 은퇴자 전략": {
        "SCHD": 0.30,
        "JEPI": 0.30,
        "BND": 0.20,
        "BIL": 0.20
    },
    "나스닥 커버드콜 극대화": {
        "JEPQ": 0.50,
        "SCHD": 0.30,
        "QQQ": 0.20
    }


}

def is_us_proxy(ticker):
    t = str(ticker).upper()
    return not t.endswith(".KS") and t != "KRW=X" and not t.isdigit()

def get_asset_name(ticker):
    # 1. Check dynamic map first (Excel-loaded assets)
    if ticker in GLOBAL_TICKER_NAME_MAP:
        return GLOBAL_TICKER_NAME_MAP[ticker]
    
    # 2. Check hardcoded universe
    for cat, assets in ASSET_UNIVERSE.items():
        if ticker in assets: return assets[ticker]
    return f"Custom({ticker})"

def fetch_hybrid_data(tickers, start_date, base_currency="KRW"):
    download_list = list(set(tickers + ["KRW=X"]))
    try:
        # auto_adjust=True를 사용하여 Close 컬럼에 수정주가를 담습니다.
        raw_data = yf.download(download_list, start=start_date, auto_adjust=True, progress=False)['Close']
        if raw_data.empty:
            raise ValueError("가져온 가격 데이터가 비어 있습니다.")
    except Exception as e:
        raise ConnectionError(f"Yahoo Finance 데이터 다운로드 실패: {str(e)}")

    # 데이터 구조가 Series인 경우(티커가 1개인 경우) DataFrame으로 변환
    if isinstance(raw_data, pd.Series):
        raw_data = raw_data.to_frame()

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
    matrix = df.groupby(['year', 'month'])['ret'].apply(lambda x: (1+x).prod()-1) * 100
    matrix = matrix.unstack()
    matrix = matrix.reindex(columns=range(1, 13))
    matrix.columns = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
    return matrix

def run_pro_backtest(port_a_w, port_b_w, start="2010-01-01", initial_investment=300_000_000, benchmark_ticker="SPY", rebalance="Monthly", base_currency="KRW", monthly_contribution=0, transaction_fee=0.0, rebalance_threshold=0.0):
    all_tickers = list(set(list(port_a_w.keys()) + list(port_b_w.keys()) + [benchmark_ticker]))
    returns, raw_monthly_prices = fetch_hybrid_data(all_tickers, start, base_currency)
    
    avail_tickers = returns.columns.tolist()
    
    def calc_port_values(w_dict, rebalance_type, fee_rate, threshold):
        valid_w = {t: v for t, v in w_dict.items() if t in avail_tickers}
        if not valid_w: 
            return pd.Series(float(initial_investment), index=returns.index), \
                   pd.Series(float(initial_investment), index=returns.index), \
                   pd.Series(0.0, index=avail_tickers)
        
        target_w = pd.Series(0.0, index=avail_tickers)
        for t, val in valid_w.items(): target_w[t] = val
        if target_w.sum() > 0: target_w /= target_w.sum()
        
        # 초기 매수 수수료 반영
        current_value = float(initial_investment) * (1 - fee_rate)
        total_invested = float(initial_investment)
        values = []
        invested_history = []
        
        contrib_history = pd.DataFrame(0.0, index=returns.index, columns=avail_tickers)
        
        # 리밸런싱 주기 설정
        step = 1 if rebalance_type == "Monthly" else (3 if rebalance_type == "Quarterly" else (12 if rebalance_type == "Yearly" else 999999))
        if rebalance_type == "Threshold":
            step = 999999 # 주기적 리밸런싱은 끄고 임계치만 체크
            
        curr_w_vec = target_w.values
        
        for i in range(len(returns)):
            ret_row = returns.iloc[i].values
            
            # 1. 수익률 반영 전 가치 (기여도 계산용)
            contrib_row = curr_w_vec * ret_row
            contrib_history.iloc[i] = contrib_row
            
            # 2. 자산 가치 업데이트 (수익률 반영)
            current_value = np.dot(curr_w_vec * current_value, (1 + ret_row))
            
            # 3. 월 적립금 투입 (적립 시에도 수수료 발생 가정)
            current_value += monthly_contribution * (1 - fee_rate)
            total_invested += monthly_contribution
            
            # 4. 현재 비중 업데이트
            curr_w_vec = curr_w_vec * (1 + ret_row)
            if curr_w_vec.sum() > 0: 
                curr_w_vec /= curr_w_vec.sum()
            
            # 5. 리밸런싱 여부 판단
            do_rebalance = False
            if (i + 1) % step == 0:
                do_rebalance = True
            elif rebalance_type == "Threshold" and threshold > 0:
                # 목표 비중과 현재 비중의 차이 중 최대값이 임계치를 넘으면 리밸런싱
                diff = np.abs(curr_w_vec - target_w.values)
                if np.any(diff > threshold):
                    do_rebalance = True
            
            # 6. 리밸런싱 실행
            if do_rebalance:
                # 매매 비용 계산: |현재비중 - 목표비중| * 전체가치 * 수수료율
                trade_amount = np.sum(np.abs(curr_w_vec - target_w.values)) * current_value
                current_value -= (trade_amount / 2) * fee_rate # 매수/매도 합산 비용의 절반(대략적) 또는 전체 적용
                curr_w_vec = target_w.values
            
            values.append(current_value)
            invested_history.append(total_invested)
            
        return pd.Series(values, index=returns.index), \
               pd.Series(invested_history, index=returns.index), \
               contrib_history.sum()

    # 포트폴리오 가치 및 기여도 계산
    has_a = bool(port_a_w)
    has_b = bool(port_b_w)

    if has_a:
        val_a, inv_a, con_a = calc_port_values(port_a_w, rebalance, transaction_fee, rebalance_threshold)
    else:
        val_a, inv_a, con_a = pd.Series(0.0, index=returns.index), pd.Series(0.0, index=returns.index), pd.Series(0.0, index=avail_tickers)

    if has_b:
        val_b, inv_b, con_b = calc_port_values(port_b_w, rebalance, transaction_fee, rebalance_threshold)
    else:
        val_b, inv_b, con_b = pd.Series(0.0, index=returns.index), pd.Series(0.0, index=returns.index), pd.Series(0.0, index=avail_tickers)
        
    val_bench, inv_bench, _ = calc_port_values({benchmark_ticker: 1.0}, "None", 0, 0)

    def get_ext_metrics(val_series, inv_series, rf=0.0):
        if val_series.empty or val_series.sum() == 0: 
            return {k: "-" for k in ["총 투입 원금(원)", "최종 평가 금액(원)", "누적 수익률(%)", "CAGR(%)", "연변동성(%)", "샤프비율", "MDD(%)"]}
        
        final_val = val_series.iloc[-1]
        total_inv = inv_series.iloc[-1]
        total_ret = (final_val / total_inv) - 1
        years = len(val_series) / 12
        cagr = (final_val / total_inv) ** (1/years) - 1 if years > 0 else 0
        
        # 수익률 기반 변동성 및 샤프비율 계산
        rets = val_series.pct_change().dropna()
        vol = rets.std() * np.sqrt(12)
        sharpe = (cagr - rf) / vol if vol > 0 else 0
        mdd = (val_series / val_series.cummax() - 1).min()
        
        return {
            "총 투입 원금(원)": f"{total_inv:,.0f}",
            "최종 평가 금액(원)": f"{final_val:,.0f}",
            "누적 수익률(%)": f"{total_ret*100:.2f}",
            "CAGR(%)": f"{cagr*100:.2f}",
            "연변동성(%)": f"{vol*100:.2f}",
            "샤프비율": f"{sharpe:.2f}",
            "MDD(%)": f"{mdd*100:.2f}"
        }

    metrics = pd.DataFrame({
        'Port A': get_ext_metrics(val_a, inv_a),
        'Port B': get_ext_metrics(val_b, inv_b),
        'Benchmark': get_ext_metrics(val_bench, inv_bench)
    })

    def get_corr(w_dict):
        if not w_dict: return pd.DataFrame()
        tickers = [t for t in w_dict.keys() if t in avail_tickers]
        if len(tickers) < 2: return pd.DataFrame()
        corr = returns[tickers].corr()
        corr.index = [get_asset_name(t).split(' (')[0] for t in corr.index]
        corr.columns = [get_asset_name(t).split(' (')[0] for t in corr.columns]
        return corr

    return {
        "asset_values_a": val_a if has_a else None, 
        "asset_values_b": val_b if has_b else None, 
        "asset_values_bench": val_bench,
        "invested_a": inv_a, "invested_b": inv_b,
        "metrics": metrics,
        "monthly_a": get_monthly_matrix(val_a.pct_change().fillna(0)) if has_a else pd.DataFrame(), 
        "monthly_b": get_monthly_matrix(val_b.pct_change().fillna(0)) if has_b else pd.DataFrame(),
        "corr_a": get_corr(port_a_w), "corr_b": get_corr(port_b_w),
        "contrib_a": con_a if has_a else None, 
        "contrib_b": con_b if has_b else None,
        "drawdown_a": (val_a / val_a.cummax() - 1) * 100 if has_a else None,
        "drawdown_b": (val_b / val_b.cummax() - 1) * 100 if has_b else None,
        "drawdown_bench": (val_bench / val_bench.cummax() - 1) * 100,
        "raw_returns": returns,
        "raw_prices": raw_monthly_prices
    }
