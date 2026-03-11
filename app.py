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
st.set_page_config(page_title="K-Global 자산배분 SaaS", layout="wide")

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
        st.title("🔐 계정 관리")
        choice = st.radio("모드 선택", ["로그인", "회원가입"])
        email = st.text_input("이메일")
        password = st.text_input("비밀번호", type="password")

        if choice == "로그인":
            if st.button("로그인"):
                try:
                    res = supabase.auth.sign_in_with_password({"email": email, "password": password})
                    st.session_state.user = res.user
                    st.rerun()
                except:
                    st.error("로그인 실패: 정보를 확인하세요.")
        else:
            if st.button("회원가입"):
                try:
                    supabase.auth.sign_up({"email": email, "password": password})
                    st.success("회원가입 성공! 이제 로그인해 주세요.")
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
            selected = st.sidebar.selectbox("저장된 전략 불러오기", ["선택하세요..."] + strat_names)
            
            if selected != "선택하세요...":
                col_load, col_del = st.sidebar.columns(2)
                if col_load.button("불러오기"):
                    data = next(s for s in strategies if s['name'] == selected)
                    st.session_state.port_a = data['port_a']
                    st.session_state.port_b = data['port_b']
                    # config 반영 (선택 사항)
                    st.sidebar.success(f"'{selected}' 로드 완료!")
                    st.rerun()
                if col_del.button("삭제"):
                    supabase.table("portfolios").delete().eq("name", selected).execute()
                    st.sidebar.warning("전략 삭제됨.")
                    st.rerun()
        else:
            st.sidebar.info("저장된 전략이 없습니다.")
    except: pass

    st.sidebar.markdown("---")
    new_name = st.sidebar.text_input("현재 전략 저장하기 (이름)")
    if st.sidebar.button("💾 클라우드 저장"):
        if new_name:
            try:
                payload = {
                    "user_id": st.session_state.user.id,
                    "name": new_name,
                    "config": {"bench": st.session_state.get('bench', 'SPY'), "cur": st.session_state.get('cur', 'KRW')},
                    "port_a": st.session_state.port_a,
                    "port_b": st.session_state.port_b
                }
                supabase.table("portfolios").insert(payload).execute()
                st.sidebar.success("저장 완료!")
                st.rerun()
            except Exception as e: st.sidebar.error(f"저장 실패: {str(e)}")
        else: st.sidebar.error("이름을 입력하세요.")

# --- [Main] 메인 애플리케이션 로직 ---
if not st.session_state.user:
    st.title("📈 K-Global 하이브리드 자산배분 서비스")
    st.markdown("### 전문가용 자산배분 엔진을 웹에서 경험해 보세요.")
    st.info("로그인 후 서비스를 이용하실 수 있습니다.")
    auth_ui()
else:
    # 1. 사이드바 설정
    st.sidebar.title(f"👤 {st.session_state.user.email}")
    if st.sidebar.button("로그아웃"):
        st.session_state.user = None
        st.rerun()
        
    strategy_library_ui()
    
    st.sidebar.markdown("---")
    st.sidebar.header("⚙️ 시뮬레이션 설정")
    start_date = st.sidebar.date_input("시작일", datetime(2005, 1, 1))
    initial_investment = st.sidebar.number_input("초기 투자금", value=300_000_000, step=10_000_000)
    rebalance_freq = st.sidebar.selectbox("리밸런싱 주기", ["Monthly", "Quarterly", "Yearly", "None"])
    base_currency = st.sidebar.selectbox("기준 통화", ["KRW", "USD"])
    benchmark = st.sidebar.selectbox("벤치마크 지수", ["SPY", "QQQ", "EWY", "VTI", "069500.KS"])

    # 2. 메인 화면: 포트폴리오 설계
    st.title("📈 K-Global 하이브리드 자산배분 (SaaS v1.0)")
    
    # 전략 프리셋
    preset = st.selectbox("전략 프리셋 적용", ["선택 안함"] + list(backtest_proxy.STRATEGY_PRESETS.keys()))
    if preset != "선택 안함":
        if st.button("프리셋 즉시 반영"):
            st.session_state.port_a = backtest_proxy.STRATEGY_PRESETS[preset]
            st.session_state.port_b = backtest_proxy.STRATEGY_PRESETS[preset]
            st.rerun()

    col1, col2 = st.columns(2)

    def portfolio_editor(which, column):
        with column:
            st.subheader(f"Portfolio {which}")
            
            # 유니버스 선택
            universe = []
            for cat, assets in backtest_proxy.ASSET_UNIVERSE.items():
                for t, n in assets.items(): universe.append(f"{t} | {n.split(' (')[0]}")
            
            defaults = [f"{t} | {backtest_proxy.get_asset_name(t).split(' (')[0]}" for t in (st.session_state.port_a if which=='A' else st.session_state.port_b).keys()]
            selected = st.multiselect(f"자산 선택 ({which})", options=universe, default=defaults, key=f"sel_{which}")
            
            # 비중 편집
            new_weights = {}
            if selected:
                st.write("**비중 설정 (%)**")
                cols = st.columns([2, 1])
                for asset_str in selected:
                    ticker = asset_str.split(" | ")[0]
                    # 세션 상태값 우선 순위
                    cur_w = (st.session_state.port_a if which=='A' else st.session_state.port_b).get(ticker, 0.0)
                    w = st.number_input(f"{asset_str}", value=float(cur_w*100), step=5.0, key=f"w_{which}_{ticker}", format="%.1f")
                    new_weights[ticker] = w / 100.0
                
                # 상태 업데이트
                if which == 'A': st.session_state.port_a = new_weights
                else: st.session_state.port_b = new_weights
                
                total = sum(new_weights.values()) * 100
                st.progress(min(total/100.0, 1.0))
                st.write(f"합계 비중: **{total:.1f}%**")
            return new_weights

    weights_a = portfolio_editor("A", col1)
    weights_b = portfolio_editor("B", col2)

    st.markdown("---")

    # 3. 백테스트 실행 및 결과 시각화
    if st.button("🚀 전체 백테스트 실행 (20년 장기 분석)", use_container_width=True):
        if not weights_a or not weights_b:
            st.error("두 포트폴리오 모두 자산을 선택해 주세요.")
        elif abs(sum(weights_a.values()) - 1.0) > 0.05 or abs(sum(weights_b.values()) - 1.0) > 0.05:
            st.error("비중 합계가 100% (±5%)여야 합니다.")
        else:
            try:
                with st.spinner("미국 Proxy 엔진을 사용하여 장기 성과를 분석 중입니다..."):
                    res = backtest_proxy.run_pro_backtest(
                        weights_a, weights_b,
                        start=start_date.strftime("%Y-%m-%d"),
                        initial_investment=initial_investment,
                        benchmark_ticker=benchmark,
                        rebalance=rebalance_freq,
                        base_currency=base_currency
                    )
                    
                    # --- [결과 1] 성과 지표 ---
                    st.header("📊 성과 지표 비교")
                    st.table(res['metrics'])
                    
                    # --- [결과 2] 자산 가치 & 드로우다운 (대형 차트) ---
                    st.header("📈 성장 및 리스크 분석")
                    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 12))
                    
                    # Equity Curve
                    ax1.plot(res['asset_values_a'], label="Portfolio A", lw=3, color='#3498DB')
                    ax1.plot(res['asset_values_b'], label="Portfolio B", lw=3, color='#E67E22')
                    ax1.plot(res['asset_values_bench'], label=f"Benchmark ({benchmark})", ls="--", color="gray", alpha=0.7)
                    ax1.set_title(f"Cumulative Equity Growth ({base_currency} Base)", fontsize=15, pad=20, fontweight='bold')
                    ax1.legend(); ax1.grid(True, ls=":", alpha=0.6)
                    ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: format(int(x), ',')))
                    
                    # Drawdown
                    ax2.fill_between(res['drawdown_a'].index, res['drawdown_a'], color='#3498DB', alpha=0.15)
                    ax2.plot(res['drawdown_a'], label="Port A DD", color='#3498DB')
                    ax2.plot(res['drawdown_b'], label="Port B DD", color='#E67E22')
                    ax2.plot(res['drawdown_bench'], label="Bench DD", ls="--", color="gray", alpha=0.6)
                    ax2.set_title("Drawdown Analysis (%)", fontsize=15, pad=20, fontweight='bold')
                    ax2.legend(); ax2.grid(True, ls=":", alpha=0.6)
                    
                    plt.tight_layout(pad=5.0)
                    st.pyplot(fig)
                    
                    # --- [결과 3] 인사이트 탭 (상관관계 & 월별) ---
                    st.header("🔗 정밀 분석 인사이트")
                    tab_corr, tab_month = st.tabs(["Asset Correlation", "Monthly Returns Matrix"])
                    
                    with tab_corr:
                        st.subheader("자산 간 상관관계 히트맵")
                        c_fig, (c_ax1, c_ax2) = plt.subplots(2, 1, figsize=(14, 16))
                        
                        def draw_st_corr(ax, df, title):
                            if not df.empty:
                                im = ax.imshow(df.values, cmap='RdYlGn', vmin=-1, vmax=1, aspect='auto')
                                ax.set_title(title, pad=25, fontsize=14, fontweight='bold')
                                ax.set_xticks(range(len(df.columns))); ax.set_yticks(range(len(df.index)))
                                ax.set_xticklabels(df.columns, rotation=45, ha='right', fontsize=10)
                                ax.set_yticklabels(df.index, fontsize=10)
                                for i in range(len(df.index)):
                                    for j in range(len(df.columns)):
                                        ax.text(j, i, f"{df.iloc[i,j]:.2f}", ha="center", va="center", fontweight='bold')
                                c_fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
                        
                        draw_st_corr(c_ax1, res['corr_a'], "Portfolio A Asset Correlation")
                        draw_st_corr(c_ax2, res['corr_b'], "Portfolio B Asset Correlation")
                        plt.tight_layout(pad=7.0)
                        st.pyplot(c_fig)
                        
                    with tab_month:
                        st.subheader("월별 수익률 히트맵")
                        m_fig, (m_ax1, m_ax2) = plt.subplots(2, 1, figsize=(16, 18))
                        
                        def draw_st_month(ax, df, title):
                            if not df.empty:
                                im = ax.imshow(df.values, cmap='RdYlGn', vmin=-5, vmax=5, aspect='auto')
                                ax.set_title(title, pad=25, fontsize=14, fontweight='bold')
                                ax.set_yticks(range(len(df.index))); ax.set_yticklabels(df.index, fontsize=10)
                                ax.set_xticks(range(len(df.columns))); ax.set_xticklabels([f"{m}M" for m in df.columns], fontsize=10)
                                for i in range(len(df.index)):
                                    for j in range(len(df.columns)):
                                        val = df.iloc[i, j]
                                        if pd.notna(val):
                                            txt = ax.annotate(f"{float(val):.1f}", xy=(j, i), ha="center", va="center", color="black", fontsize=9, fontweight='bold')
                                            txt.set_path_effects([patheffects.withStroke(linewidth=2, foreground='white')])
                                m_fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
                        
                        draw_st_month(m_ax1, res['monthly_a'], "Portfolio A Monthly Returns (%)")
                        draw_st_month(m_ax2, res['monthly_b'], "Portfolio B Monthly Returns (%)")
                        plt.tight_layout(pad=8.0)
                        st.pyplot(m_fig)

                    # --- [결과 4] 상세 엑셀 보고서 다운로드 ---
                    st.header("📥 보고서 데이터 내보내기")
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        res['metrics'].to_excel(writer, sheet_name='Summary')
                        # 상세 가격 데이터 (1억 투자 시뮬레이션)
                        for p_idx, p_data in [('A', weights_a), ('B', weights_b)]:
                            valid_t = [t for t in p_data.keys() if t in res['raw_returns'].columns]
                            if valid_t:
                                fx_data = res['raw_prices']['KRW=X'] if 'KRW=X' in res['raw_prices'].columns else 1.0
                                detail_dfs = []
                                for t in valid_t:
                                    t_df = pd.DataFrame(index=res['raw_returns'].index)
                                    t_df[f'{t}_Market_Price'] = res['raw_prices'][t]
                                    t_df['Exchange_Rate'] = fx_data
                                    t_df[f'{t}_Value_100M'] = (1 + res['raw_returns'][t]).cumprod() * 100_000_000
                                    detail_dfs.append(t_df)
                                pd.concat(detail_dfs, axis=1).to_excel(writer, sheet_name=f'Price_Detail_{p_idx}')
                        res['monthly_a'].to_excel(writer, sheet_name='Monthly_A')
                        res['monthly_b'].to_excel(writer, sheet_name='Monthly_B')
                        res['corr_a'].to_excel(writer, sheet_name='Corr_A')
                        res['corr_b'].to_excel(writer, sheet_name='Corr_B')
                    
                    st.download_button(
                        label="📥 전문가용 상세 엑셀 보고서 다운로드",
                        data=output.getvalue(),
                        file_name=f"K-Global_Full_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            except Exception as e:
                st.error(f"백테스트 분석 중 오류가 발생했습니다: {str(e)}")

st.sidebar.markdown("---")
st.sidebar.info("Professional Asset Allocation SaaS v1.0")
