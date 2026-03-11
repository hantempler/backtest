import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import backtest_proxy
from datetime import datetime
import io

# --- 페이지 설정 ---
st.set_page_config(page_title="K-Global 자산배분 시뮬레이터", layout="wide")

# --- Matplotlib 한글 폰트 설정 ---
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

# --- 세션 상태 초기화 ---
if 'port_a' not in st.session_state:
    st.session_state.port_a = {}
if 'port_b' not in st.session_state:
    st.session_state.port_b = {}

# --- 사이드바: 설정 ---
st.sidebar.header("⚙️ Backtest Configuration")
start_date = st.sidebar.date_input("Start Date", datetime(2005, 1, 1))
initial_investment = st.sidebar.number_input("Initial Investment (KRW/USD)", value=300_000_000, step=10_000_000)
rebalance_freq = st.sidebar.selectbox("Rebalancing Frequency", ["Monthly", "Quarterly", "Yearly", "None"])
base_currency = st.sidebar.selectbox("Base Currency", ["KRW", "USD"])
benchmark = st.sidebar.selectbox("Benchmark", ["SPY", "QQQ", "EWY", "VTI", "069500.KS"])

# --- 메인 화면 ---
st.title("📈 K-Global 하이브리드 자산배분 (Web v1.0)")
st.markdown("---")

# 자산 선택 및 비중 설정
col1, col2 = st.columns(2)

def portfolio_ui(which, column):
    with column:
        st.subheader(f"Portfolio {which}")
        
        # 유니버스 카테고리별 선택
        all_assets = []
        for cat, assets in backtest_proxy.ASSET_UNIVERSE.items():
            for t, n in assets.items():
                all_assets.append(f"{t} | {n}")
        
        selected_assets = st.multiselect(
            f"Select Assets for {which}",
            options=all_assets,
            default=[f"{t} | {backtest_proxy.get_asset_name(t)}" for t in (st.session_state.port_a if which == 'A' else st.session_state.port_b).keys()],
            key=f"select_{which}"
        )
        
        # 비중 입력
        weights = {}
        if selected_assets:
            st.write("Set Weights (%)")
            total_selected = len(selected_assets)
            eq_w = 100.0 / total_selected
            
            for asset_str in selected_assets:
                ticker = asset_str.split(" | ")[0]
                w = st.number_input(f"{asset_str}", value=eq_w, key=f"w_{which}_{ticker}")
                weights[ticker] = w / 100.0
            
            sum_w = sum(weights.values()) * 100
            st.info(f"Current Sum: {sum_w:.1f}%")
            if abs(sum_w - 100.0) > 0.1:
                st.warning("Total weight must be 100%.")
        
        return weights

weights_a = portfolio_ui("A", col1)
weights_b = portfolio_ui("B", col2)

st.markdown("---")

# --- 실행 버튼 ---
if st.button("🚀 RUN BACKTEST", use_container_width=True):
    if not weights_a or not weights_b:
        st.error("Please select assets for both portfolios.")
    elif abs(sum(weights_a.values()) - 1.0) > 0.01 or abs(sum(weights_b.values()) - 1.0) > 0.01:
        st.error("Weights for both portfolios must sum to 100%.")
    else:
        try:
            with st.spinner("Analyzing global markets..."):
                res = backtest_proxy.run_pro_backtest(
                    weights_a, weights_b,
                    start=start_date.strftime("%Y-%m-%d"),
                    initial_investment=initial_investment,
                    benchmark_ticker=benchmark,
                    rebalance=rebalance_freq,
                    base_currency=base_currency
                )
                
                # 1. 성과 요약
                st.header("📊 Performance Summary")
                st.dataframe(res['metrics'], use_container_width=True)
                
                # 2. 자산 가치 & 드로우다운
                st.header("📈 Growth & Risk Analysis")
                fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10))
                
                ax1.plot(res['asset_values_a'], label="Port A", lw=2)
                ax1.plot(res['asset_values_b'], label="Port B", lw=2)
                ax1.plot(res['asset_values_bench'], label=f"Bench({benchmark})", ls="--", color="gray")
                ax1.set_title(f"Cumulative Equity Growth ({base_currency})")
                ax1.legend(); ax1.grid(True, ls=":")
                
                ax2.fill_between(res['drawdown_a'].index, res['drawdown_a'], alpha=0.1)
                ax2.plot(res['drawdown_a'], label="A DD")
                ax2.plot(res['drawdown_b'], label="B DD")
                ax2.plot(res['drawdown_bench'], label="Bench DD", ls="--", color="gray")
                ax2.set_title("Drawdown (%)")
                ax2.legend(); ax2.grid(True, ls=":")
                
                st.pyplot(fig)
                
                # 3. 상관관계 & 월별 수익률 (히트맵)
                st.header("🔗 Insights")
                tab1, tab2 = st.tabs(["Asset Correlation", "Monthly Returns"])
                
                with tab1:
                    c_col1, c_col2 = st.columns(2)
                    c_col1.subheader("Portfolio A")
                    c_col1.dataframe(res['corr_a'].style.background_gradient(cmap='RdYlGn', vmin=-1, vmax=1))
                    c_col2.subheader("Portfolio B")
                    c_col2.dataframe(res['corr_b'].style.background_gradient(cmap='RdYlGn', vmin=-1, vmax=1))
                
                with tab2:
                    st.subheader("Monthly Matrix (A)")
                    st.dataframe(res['monthly_a'].style.background_gradient(cmap='RdYlGn', vmin=-5, vmax=5).format("{:.1f}%"))
                    st.subheader("Monthly Matrix (B)")
                    st.dataframe(res['monthly_b'].style.background_gradient(cmap='RdYlGn', vmin=-5, vmax=5).format("{:.1f}%"))

                # 엑셀 다운로드
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    res['metrics'].to_excel(writer, sheet_name='Summary')
                    res['monthly_a'].to_excel(writer, sheet_name='Monthly_A')
                    res['monthly_b'].to_excel(writer, sheet_name='Monthly_B')
                
                st.download_button(
                    label="📥 Download Detailed Excel Report",
                    data=output.getvalue(),
                    file_name=f"backtest_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
        except Exception as e:
            st.error(f"Analysis Failed: {str(e)}")

st.sidebar.markdown("---")
st.sidebar.info("Professional K-Global Backtest Engine")
