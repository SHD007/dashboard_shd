# app.py
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ---------- PAGE / THEME ----------
st.set_page_config(page_title="시각화 대시보드", page_icon="📊", layout="wide")

st.markdown("""
<style>
/* 카드 느낌 */
.kpi {background:rgba(255,255,255,.03); border:1px solid rgba(255,255,255,.08);
      padding:14px 16px; border-radius:16px}
[data-theme="light"] .kpi {background:#fff; border-color:#eee}
.small {opacity:.8; font-size:.9rem}
.block {margin-bottom:1rem}
h1, h2, h3 {letter-spacing:.2px}
</style>
""", unsafe_allow_html=True)

# ---------- SIDEBAR ----------
st.sidebar.header("설정")
theme = st.sidebar.selectbox("Plotly 테마", ["plotly_dark", "plotly_white"], index=0)
palette = st.sidebar.selectbox(
    "색상 팔레트",
    ["Plotly", "Pastel", "Dark24", "Set2", "Set3"],
    index=2
)
palette_map = {
    "Plotly": px.colors.qualitative.Plotly,
    "Pastel": px.colors.qualitative.Pastel,
    "Dark24": px.colors.qualitative.Dark24,
    "Set2": px.colors.qualitative.Set2,
    "Set3": px.colors.qualitative.Set3,
}
colorway = palette_map[palette]

uploaded = st.sidebar.file_uploader("엑셀 업로드 (.xlsx)", type=["xlsx"])
st.sidebar.caption("시트명: 바차트_히스토그램 / 시계열차트 / 파이차트 / 산점도 / 파레토차트 / 버블차트")

@st.cache_data(show_spinner=False)
def load_book(file) -> dict:
    xls = pd.ExcelFile(file)
    return {sn: pd.read_excel(xls, sn) for sn in xls.sheet_names}

def to_month(df, col="월"):
    if col in df.columns:
        df = df.copy()
        df[col] = pd.to_datetime(df[col], errors="coerce")
    return df

# 샘플 데이터(업로드 없을 때 테스트 용)
def sample():
    dates = pd.date_range("2023-01-01", periods=12, freq="M")
    a = [272,147,217,292,423,259,216,370,315,481,299,372]
    b = [118,109,155,177,170,162,180,171,201,205,212,195]
    c = [276,206,204,279,307,248,217,169,260,302,242,288]
    d = [145,185,153,200,198,212,152,193,176,174,223,198]
    e = [74,145,158,200,338,154,162,117,150,106,215,183]
    total = [x+y+z+w+v for x,y,z,w,v in zip(a,b,c,d,e)]
    return {
        "바차트_히스토그램": pd.DataFrame({"월": dates, "총 매출": total}),
        "시계열차트": pd.DataFrame({"월": dates, "제품 A 매출": a, "제품 B 매출": b,
                                  "제품 C 매출": c, "제품 D 매출": d, "제품 E 매출": e}),
        "파이차트": pd.DataFrame({"제품": ["제품 A","제품 B","제품 C","제품 D","제품 E"],
                               "1분기 매출":[sum(a[:3]),sum(b[:3]),sum(c[:3]),sum(d[:3]),sum(e[:3])] }),
        "산점도": pd.DataFrame({"제품 A 매출": a, "비용":[149,227,293,335,197,255,190,240,265,310,205,280]}),
        "파레토차트": pd.DataFrame({"부서":["기획부","마케팅부","영업부","인사부","개발부"],
                                 "매출":[954,923,559,477,209]}),
        "버블차트": pd.DataFrame({"제품":[f"제품 {i}" for i in range(1,11)],
                               "제품별 비용":[884,759,829,392,963,510,610,420,700,560],
                               "마진":[699,170,572,496,414,380,420,210,460,330],
                               "고객 수":[127,122,59,198,165,110,130,95,140,120]})
    }

if uploaded:
    dfs = load_book(uploaded)
else:
    dfs = sample()

# ---------- DATA ----------
bar_df     = to_month(dfs.get("바차트_히스토그램", pd.DataFrame()))
series_df  = to_month(dfs.get("시계열차트", pd.DataFrame()))
pie_df     = dfs.get("파이차트", pd.DataFrame())
scatter_df = dfs.get("산점도", pd.DataFrame())
pareto_df  = dfs.get("파레토차트", pd.DataFrame())

# ---------- HEADER ----------
st.title("📊 시각화 대시보드")
st.caption("엑셀 파일(.xlsx) 업로드로 6개 차트를 자동 생성합니다. 다크 모드에 최적화된 팔레트를 사용합니다.")

# ---------- KPI CARDS ----------
k1, k2, k3, k4 = st.columns(4)
with k1:
    if not bar_df.empty and {"월","총 매출"}.issubset(bar_df.columns):
        st.markdown('<div class="kpi">', unsafe_allow_html=True)
        st.metric("연간 총 매출", f"{bar_df['총 매출'].sum():,}")
        st.markdown('<div class="small">바차트 시트 기준</div></div>', unsafe_allow_html=True)
with k2:
    if not series_df.empty and "월" in series_df.columns and len(series_df.columns)>1:
        last = series_df.sort_values("월").tail(1)
        cols = [c for c in series_df.columns if c!="월"]
        st.markdown('<div class="kpi">', unsafe_allow_html=True)
        st.metric("최근월 총 매출", f"{int(last[cols].sum(axis=1)): ,}")
        st.markdown('</div>', unsafe_allow_html=True)
with k3:
    if not pareto_df.empty and "매출" in pareto_df.columns:
        top = pareto_df.sort_values("매출", ascending=False).iloc[0]
        st.markdown('<div class="kpi">', unsafe_allow_html=True)
        st.metric("Top 부서", top["부서"], delta=f"{int(top['매출']):,}")
        st.markdown('</div>', unsafe_allow_html=True)
with k4:
    st.markdown('<div class="kpi">', unsafe_allow_html=True)
    st.metric("시트 수", len(dfs))
    st.markdown('</div>', unsafe_allow_html=True)

# ---------- ROW 1: Bar & Line ----------
c1, c2 = st.columns(2, gap="large")
with c1:
    st.subheader("월별 총 매출")
    if not bar_df.empty and {"월","총 매출"}.issubset(bar_df.columns):
        bar_df = bar_df.sort_values("월")
        bar_df["월_라벨"] = bar_df["월"].dt.strftime("%Y-%m")
        fig = px.bar(bar_df, x="월_라벨", y="총 매출",
                     template=theme, color_discrete_sequence=[colorway[3]])
        fig.update_traces(hovertemplate="월=%{x}<br>매출=%{y:,}")
        st.plotly_chart(fig, use_container_width=True)
with c2:
    st.subheader("시계열 추세")
    if not series_df.empty and "월" in series_df.columns and len(series_df.columns)>1:
        series_df = series_df.sort_values("월")
        melt = series_df.melt(id_vars="월", var_name="항목", value_name="값")
        fig = px.line(melt, x="월", y="값", color="항목",
                      markers=True, template=theme, color_discrete_sequence=colorway)
        fig.update_traces(hovertemplate="%{x|%Y-%m}<br>%{legendgroup}=%{y:,}")
        st.plotly_chart(fig, use_container_width=True)

# ---------- ROW 2: Pie & Scatter ----------
c3, c4 = st.columns(2, gap="large")
with c3:
    st.subheader("비율 분석")
    if not pie_df.empty:
        label_col = pie_df.columns[0]
        value_col = "1분기 매출" if "1분기 매출" in pie_df.columns else \
            (pie_df.select_dtypes(include="number").columns[:1].tolist() or [None])[0]
        if value_col:
            fig = px.pie(pie_df, names=label_col, values=value_col,
                         template=theme, color_discrete_sequence=colorway, hole=0)
            fig.update_traces(textposition="inside", textinfo="percent+label")
            st.plotly_chart(fig, use_container_width=True)
with c4:
    st.subheader("산점도 분석")
    if not scatter_df.empty and {"제품 A 매출","비용"}.issubset(scatter_df.columns):
        fig = px.scatter(scatter_df, x="제품 A 매출", y="비용",
                         template=theme, color_discrete_sequence=[colorway[4]],
                         hover_data=scatter_df.columns)
        st.plotly_chart(fig, use_container_width=True)

# ---------- ROW 3: Pareto ----------
st.subheader("파레토 차트")
if not pareto_df.empty and {"부서","매출"}.issubset(pareto_df.columns):
    p = pareto_df.sort_values("매출", ascending=False).copy()
    p["누적비율(%)"] = p["매출"].cumsum() / p["매출"].sum() * 100
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_bar(x=p["부서"], y=p["매출"], name="매출",
                marker_color=colorway[3])
    fig.add_trace(go.Scatter(x=p["부서"], y=p["누적비율(%)"],
                             name="누적 비율(%)", mode="lines+markers"),
                  secondary_y=True)
    fig.update_layout(template=theme)
    fig.update_yaxes(title_text="매출", secondary_y=False)
    fig.update_yaxes(title_text="누적 비율(%)", range=[0,110], secondary_y=True)
    st.plotly_chart(fig, use_container_width=True)

# ---------- RAW DATA ----------
with st.expander("원시 데이터 미리보기"):
    for name, df in dfs.items():
        st.markdown(f"**{name}**")
        st.dataframe(df, use_container_width=True)
