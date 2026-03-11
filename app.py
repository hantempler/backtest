import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import backtest_proxy
from datetime import datetime
import io
import os
from dotenv import load_dotenv
from supabase import create_client, Client
from matplotlib import patheffects

# --- 설정 및 환경 변수 로드 ---
load_dotenv()
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")

if not SUPABASE_URL or not SUPABASE_KEY:
    st.error(".env 파일에 SUPABASE_URL과 SUPABASE_KEY를 설정해주세요.")
    st.stop()

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# --- 페이지 설정 ---
st.set_page_config(page_title="K-Global 자산배분 SaaS Pro", layout="wide")

# --- Matplotlib 한글 폰트 설정 ---
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

# --- 세션 상태 관리 ---
if 'user' not in st.session_state:
    st.session_state.user = None
if 'port_a' not in st.session_state:
    st.session_state.port_a = {}
if 'port_b' not in st.session_state:
    st.session_state.port_b = {}

# --- [Auth] 로그인/회원가입 UI ---
def auth_ui():
    with st.sidebar:
        st.title("🔐 서비스 계정 관리")
        choice = st.radio("모드 선택", ["로그인", "회원가입"])
        email = st.text_input("이메일")
        password = st.text_input("비밀번호", type="password")

        if choice == "로그인":
            if st.button("로그인", use_container_width=True):
                try:
                    res = supabase.auth.sign_in_with_password({"email": email, "password": password})
                    st.session_state.user = res.user
                    st.rerun()
                except:
                    st.error("로그인 실패: 정보를 확인하세요.")
        else:
            if st.button("회원가입", use_container_width=True):
                try:
                    supabase.auth.sign_up({"email": email, "password": password})
                    st.success("회원가입 완료! 이제 로그인해 주세요.")
                except Exception as e:
                    st.error(f"가입 오류: {str(e)}")

# --- [Library] 전략 저장소 UI ---
def strategy_library_ui():
    if not st.session_state.user: return

    st.sidebar.markdown("---")
    st.sidebar.subheader("📁 나의 전략 보관함")
    
    try:
        res = supabase.table("portfolios").select("*").eq("user_id", st.session_state.user.id).execute()
        strategies = res.data
        if strategies:
            strat_names = [s['name'] for s in strategies]
            selected = st.sidebar.selectbox("전략 불러오기", ["선택하세요..."] + strat_names)
            
            if selected != "선택하세요...":
                col_load, col_del = st.sidebar.columns(2)
                if col_load.button("📥 로드", use_container_width=True):
                    data = next(s for s in strategies if s['name'] == selected)
                    st.session_state.port_a = data['port_a']
                    st.session_state.port_b = data['port_b']
                    st.sidebar.success(f"'{selected}' 로드 완료!")
                    st.rerun()
                if col_del.button("🗑️ 삭제", use_container_width=True):
                    supabase.table("portfolios").delete().eq("name", selected).execute()
                    st.sidebar.warning("전략 삭제됨.")
                    st.rerun()
        else:
            st.sidebar.info("저장된 전략이 없습니다.")
    except: pass

    st.sidebar.markdown("---")
    new_name = st.sidebar.text_input("현재 전략 저장하기", placeholder="전략 이름 입력")
    if st.sidebar.button("💾 클라우드 저장", use_container_width=True):
        if new_name:
            try:
                payload = {
                    "user_id": st.session_state.user.id,
                    "name": new_name,
                    "config": {
                        "bench": st.session_state.get('benchmark', 'SPY'),
                        "cur": st.session_state.get('base_currency', 'KRW'),
                        "reb": st.session_state.get('rebalance_freq', 'Monthly')
                    },
                    "port_a": st.session_state.port_a,
                    "port_b": st.session_state.port_b
                }
                supabase.table("portfolios").insert(payload).execute()
                st.sidebar.success("클라우드 저장 완료!")
                st.rerun()
            except Exception as e: st.sidebar.error(f"저장 실패: {str(e)}")
        else: st.sidebar.error("이름을 입력하세요.")

# --- [Helper] 포트폴리오 비중 초기화 ---
def reset_portfolio(which):
    if which == 'A': st.session_state.port_a = {t: 0.0 for t in st.session_state.port_a}
    else: st.session_state.port_b = {t: 0.0 for t in st.session_state.port_b}

# --- [Main] 메인 애플리케이션 시작 ---
if not st.session_state.user:
    st.title("📈 K-Global 하이브리드 자산배분 Pro")
    st.markdown("### 전문가급 자산배분 엔진을 웹에서 경험해 보세요.")
    st.info("로그인 후 서비스를 이용하실 수 있습니다.")
    auth_ui()
else:
    # 1. 사이드바 계정 및 전략 관리
    st.sidebar.title(f"👤 {st.session_state.user.email}")
    if st.sidebar.button("로그아웃", use_container_width=True):
        st.session_state.user = None
        st.rerun()
        
    strategy_library_ui()
    
    st.sidebar.markdown("---")
    st.sidebar.header("⚙️ 시뮬레이션 설정")
    start_date = st.sidebar.date_input("Backtest 시작일", datetime(2005, 1, 1))
    initial_investment = st.sidebar.number_input("초기 투자금 (원/달러)", value=300_000_000, step=10_000_000)
    rebalance_freq = st.sidebar.selectbox("리밸런싱 주기", ["Monthly", "Quarterly", "Yearly", "None"], key='rebalance_freq')
    base_currency = st.sidebar.selectbox("기준 통화 (Base Currency)", ["KRW", "USD"], key='base_currency')
    benchmark = st.sidebar.selectbox("벤치마크 지수", ["SPY", "QQQ", "EWY", "VTI", "069500.KS"], key='benchmark')

    # 2. 메인 화면: 포트폴리오 설계
    st.title("📈 Professional Asset Allocation Dashboard")
    
    # 2.1 전략 프리셋 (상단 배치)
    with st.expander("🚀 전략 프리셋 적용 (All-Weather, 60/40 등)", expanded=False):
        p_col1, p_col2 = st.columns([3, 1])
        preset_name = p_col1.selectbox("원하는 전략을 선택하세요", ["선택 안함"] + list(backtest_proxy.STRATEGY_PRESETS.keys()))
        if p_col2.button("전략 적용", use_container_width=True):
            if preset_name != "선택 안함":
                st.session_state.port_a = backtest_proxy.STRATEGY_PRESETS[preset_name].copy()
                st.session_state.port_b = backtest_proxy.STRATEGY_PRESETS[preset_name].copy()
                st.rerun()

    # 2.2 포트폴리오 섹션
    st.markdown("### 🛠️ Portfolio Configuration")
    col_a, col_b = st.columns(2)

    def portfolio_box_ui(which, column):
        with column:
            st.markdown(f"#### Portfolio {which}")
            
            # 카테고리별 자산 선택
            with st.expander(f"➕ Portfolio {which} 자산 추가", expanded=False):
                for cat, assets in backtest_proxy.ASSET_UNIVERSE.items():
                    st.write(f"**{cat}**")
                    for t, n in assets.items():
                        if st.button(f"{t} | {n.split(' (')[0]}", key=f"add_{which}_{t}"):
                            target = st.session_state.port_a if which == 'A' else st.session_state.port_b
                            if t not in target: target[t] = 0.0
                            st.rerun()
                st.markdown("---")
                c_ticker = st.text_input("커스텀 티커 입력", key=f"custom_input_{which}")
                if st.button("추가", key=f"custom_btn_{which}"):
                    if c_ticker:
                        target = st.session_state.port_a if which == 'A' else st.session_state.port_b
                        target[c_ticker.upper()] = 0.0
                        st.rerun()

            # 비중 설정 (표 형식)
            current_port = st.session_state.port_a if which == 'A' else st.session_state.port_b
            if current_port:
                st.write("**Weights (%)**")
                # 헤더
                h_col1, h_col2, h_col3 = st.columns([1, 3, 1.5])
                h_col1.caption("Ticker")
                h_col2.caption("Asset Name")
                h_col3.caption("Weight")
                
                updated_weights = {}
                for t in list(current_port.keys()):
                    r_col1, r_col2, r_col3 = st.columns([1, 3, 1.5])
                    r_col1.write(f"`{t}`")
                    name = backtest_proxy.get_asset_name(t).split(' (')[0]
                    r_col2.write(f"{name}")
                    
                    val = st.number_input(f"W_{which}_{t}", value=float(current_port[t]*100), step=1.0, key=f"input_{which}_{t}", label_visibility="collapsed", format="%.1f")
                    updated_weights[t] = val / 100.0
                    
                    if st.session_state.get(f"del_{which}_{t}"): # 삭제 로직 (사용자 요청 시)
                        pass
                
                if which == 'A': st.session_state.port_a = updated_weights
                else: st.session_state.port_b = updated_weights
                
                # 하단 조작 버튼
                b_col1, b_col2, b_col3 = st.columns(3)
                if b_col1.button(f"Reset {which}", key=f"reset_{which}"): reset_portfolio(which); st.rerun()
                if b_col2.button(f"Equal {which}", key=f"equal_{which}"):
                    n = len(updated_weights)
                    if n > 0:
                        for k in updated_weights: updated_weights[k] = 1.0/n
                        st.rerun()
                
                total = sum(updated_weights.values()) * 100
                b_col3.write(f"**Total: {total:.1f}%**")
                st.progress(min(total/100.0, 1.0))
            else:
                st.info("좌측 유니버스에서 자산을 선택하거나 커스텀 티커를 추가해 주세요.")

    portfolio_box_ui("A", col_a)
    portfolio_box_ui("B", col_b)

    st.markdown("---")

    # 3. 백테스트 실행 및 프로페셔널 리포트
    if st.button("🚀 RUN PROFESSIONAL BACKTEST", use_container_width=True):
        if not st.session_state.port_a or not st.session_state.port_b:
            st.error("두 포트폴리오 모두 최소 하나 이상의 자산이 필요합니다.")
        elif abs(sum(st.session_state.port_a.values()) - 1.0) > 0.05 or abs(sum(st.session_state.port_b.values()) - 1.0) > 0.05:
            st.error("비중 합계가 100% (±5%)여야 합니다.")
        else:
            try:
                with st.spinner("미국 Proxy 엔진을 사용하여 정밀 장기 분석을 수행 중입니다..."):
                    res = backtest_proxy.run_pro_backtest(
                        st.session_state.port_a, st.session_state.port_b,
                        start=start_date.strftime("%Y-%m-%d"),
                        initial_investment=initial_investment,
                        benchmark_ticker=benchmark,
                        rebalance=rebalance_freq,
                        base_currency=base_currency
                    )
                    
                    # --- [Part 1] Performance Overview ---
                    st.header("📊 Performance Statistics")
                    st.table(res['metrics'])
                    
                    # --- [Part 2] Main Charts ---
                    st.header("📈 Equity & Risk Visuals")
                    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 12))
                    
                    # Equity Curve (Professional Style)
                    ax1.plot(res['asset_values_a'], label="Portfolio A", lw=3, color='#3498DB')
                    ax1.plot(res['asset_values_b'], label="Portfolio B", lw=3, color='#E67E22')
                    ax1.plot(res['asset_values_bench'], label=f"Benchmark ({benchmark})", ls="--", color="#7F8C8D", alpha=0.8)
                    ax1.set_title(f"Cumulative Equity Growth ({base_currency})", fontsize=16, pad=20, fontweight='bold')
                    ax1.legend(fontsize=12); ax1.grid(True, ls=":", alpha=0.6)
                    ax1.set_yscale('linear')
                    ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format(int(x), ',')))
                    
                    # Drawdown (Professional Style)
                    ax2.fill_between(res['drawdown_a'].index, res['drawdown_a'], color='#3498DB', alpha=0.15)
                    ax2.plot(res['drawdown_a'], label="Port A DD", color='#3498DB', lw=1.5)
                    ax2.plot(res['drawdown_b'], label="Port B DD", color='#E67E22', lw=1.5)
                    ax2.plot(res['drawdown_bench'], label=f"Bench({benchmark}) DD", ls="--", color="#7F8C8D", alpha=0.7)
                    ax2.set_title("Historical Drawdown (%)", fontsize=16, pad=20, fontweight='bold')
                    ax2.legend(fontsize=12); ax2.grid(True, ls=":", alpha=0.6)
                    
                    plt.tight_layout(pad=6.0)
                    st.pyplot(fig)
                    
                    # --- [Part 3] Advanced Heatmaps ---
                    st.header("🔗 Professional Insight Matrices")
                    tab_corr, tab_month = st.tabs(["Asset Correlation Analysis", "Monthly Returns Matrix"])
                    
                    with tab_corr:
                        st.subheader("Asset Correlation (Heatmap)")
                        c_fig, (c_ax1, c_ax2) = plt.subplots(2, 1, figsize=(14, 16))
                        
                        def draw_pro_corr(ax, df, title):
                            if not df.empty:
                                im = ax.imshow(df.values, cmap='RdYlGn', vmin=-1, vmax=1, aspect='auto')
                                ax.set_title(title, pad=30, fontsize=15, fontweight='bold')
                                ax.set_xticks(range(len(df.columns))); ax.set_yticks(range(len(df.index)))
                                ax.set_xticklabels(df.columns, rotation=45, ha='right', fontsize=10)
                                ax.set_yticklabels(df.index, fontsize=10)
                                for i in range(len(df.index)):
                                    for j in range(len(df.columns)):
                                        ax.text(j, i, f"{df.iloc[i,j]:.2f}", ha="center", va="center", fontweight='bold')
                                c_fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
                        
                        draw_pro_corr(c_ax1, res['corr_a'], "Portfolio A Correlation Matrix")
                        draw_pro_corr(c_ax2, res['corr_b'], "Portfolio B Correlation Matrix")
                        plt.tight_layout(pad=8.0)
                        st.pyplot(c_fig)
                        
                    with tab_month:
                        st.subheader("Monthly Returns Distribution")
                        m_fig, (m_ax1, m_ax2) = plt.subplots(2, 1, figsize=(16, 18))
                        
                        def draw_pro_month(ax, df, title):
                            if not df.empty:
                                im = ax.imshow(df.values, cmap='RdYlGn', vmin=-5, vmax=5, aspect='auto')
                                ax.set_title(title, pad=30, fontsize=15, fontweight='bold')
                                ax.set_yticks(range(len(df.index))); ax.set_yticklabels(df.index, fontsize=10)
                                ax.set_xticks(range(len(df.columns))); ax.set_xticklabels([f"{m}M" for m in df.columns], fontsize=10)
                                for i in range(len(df.index)):
                                    for j in range(len(df.columns)):
                                        val = df.iloc[i, j]
                                        if pd.notna(val):
                                            txt = ax.annotate(f"{float(val):.1f}", xy=(j, i), ha="center", va="center", color="black", fontsize=9, fontweight='bold')
                                            txt.set_path_effects([patheffects.withStroke(linewidth=2, foreground='white')])
                                m_fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
                        
                        draw_pro_month(m_ax1, res['monthly_a'], "Portfolio A Monthly Returns (%)")
                        draw_pro_month(m_ax2, res['monthly_b'], "Portfolio B Monthly Returns (%)")
                        plt.tight_layout(pad=9.0)
                        st.pyplot(m_fig)

                    # --- [Part 4] Enhanced Excel Export (v1.4 Spec) ---
                    st.header("📥 Export Expert Report")
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        res['metrics'].to_excel(writer, sheet_name='Performance_Summary')
                        
                        # 상세 투자 가치 및 가격 데이터 (Price_Detail_A/B)
                        for p_idx, p_data in [('A', st.session_state.port_a), ('B', st.session_state.port_b)]:
                            valid_t = [t for t in p_data.keys() if t in res['raw_returns'].columns]
                            if not valid_t: continue
                            
                            detail_dfs = []
                            fx_data = res['raw_prices']['KRW=X'] if 'KRW=X' in res['raw_prices'].columns else 1.0
                            for t in valid_t:
                                t_df = pd.DataFrame(index=res['raw_returns'].index)
                                t_df[f'{t}_Market_Price'] = res['raw_prices'][t]
                                t_df['Exchange_Rate'] = fx_data
                                t_df[f'{t}_Value_100M'] = (1 + res['raw_returns'][t]).cumprod() * 100_000_000
                                detail_dfs.append(t_df)
                            pd.concat(detail_dfs, axis=1).to_excel(writer, sheet_name=f'Price_Detail_{p_idx}')
                        
                        res['monthly_a'].to_excel(writer, sheet_name='Monthly_A')
                        res['monthly_b'].to_excel(writer, sheet_name='Monthly_B')
                        res['corr_a'].to_excel(writer, sheet_name='Correlation_A')
                        res['corr_b'].to_excel(writer, sheet_name='Correlation_B')
                    
                    st.download_button(
                        label="📥 전문가용 상세 엑셀 보고서 다운로드 (Original Price & FX Included)",
                        data=output.getvalue(),
                        file_name=f"K-Global_Full_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            except Exception as e:
                st.error(f"백테스트 분석 중 오류가 발생했습니다: {str(e)}")

st.sidebar.markdown("---")
st.sidebar.info("Professional K-Global SaaS Pro v1.4")
