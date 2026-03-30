# -*- coding: utf-8 -*-
"""
코스맥스 시제품 안정성 테스트 분석 대시보드
- 시제품정보(메타) + 안정성테스트결과(본 데이터)를 조인하여 통합 분석
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

# ─────────────────────────────────────────────
# 페이지 설정
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="코스맥스 안정성 테스트 대시보드",
    page_icon="🧴",
    layout="wide",
)

st.title("🧴 코스맥스 시제품 안정성 테스트 대시보드")
st.markdown("엑셀 파일을 업로드하면 시제품정보 + 안정성테스트결과를 조인하여 분석합니다.")

# ─────────────────────────────────────────────
# 파일 업로드 & 데이터 로드
# ─────────────────────────────────────────────
uploaded_file = st.file_uploader("📁 엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

if uploaded_file is None:
    st.info("엑셀 파일을 업로드해주세요.")
    st.stop()

xls = pd.ExcelFile(uploaded_file)
sheet_names = xls.sheet_names

if len(sheet_names) < 2:
    st.error("시트가 2개 이상이어야 합니다 (시제품정보 + 안정성테스트결과)")
    st.stop()

meta = pd.read_excel(xls, sheet_name=sheet_names[0])
test = pd.read_excel(xls, sheet_name=sheet_names[1])

# 날짜 변환
for col in test.columns:
    if "일" in col or "date" in col.lower():
        test[col] = pd.to_datetime(test[col], errors="coerce")
for col in meta.columns:
    if "일" in col or "date" in col.lower():
        meta[col] = pd.to_datetime(meta[col], errors="coerce")

# 조인 (시제품코드 기준)
join_key = "시제품코드"
df = test.merge(meta, on=join_key, how="left")

st.success(
    f"✅ 로드 완료 — "
    f"{sheet_names[0]}({len(meta)}행) + {sheet_names[1]}({len(test)}행) → "
    f"조인 결과 {len(df)}행 × {len(df.columns)}열"
)

# 메타 없는 시제품 안내
no_meta = df[df[meta.columns.difference([join_key])].isnull().all(axis=1)][join_key].unique()
if len(no_meta) > 0:
    st.warning(f"⚠️ 메타정보 없는 시제품: {', '.join(no_meta)} — 테스트 데이터만 분석됩니다.")

# ─────────────────────────────────────────────
# 사이드바 필터
# ─────────────────────────────────────────────
st.sidebar.header("🔍 필터")

filtered = df.copy()

filter_cols = ["시제품코드", "테스트조건", "판정결과", "제품유형", "개발단계", "목표피부타입", "주요컨셉", "담당팀"]
for col in filter_cols:
    if col in df.columns:
        unique_vals = sorted(df[col].dropna().unique().tolist())
        if 0 < len(unique_vals) <= 20:
            options = ["전체"] + unique_vals
            selected = st.sidebar.selectbox(col, options, index=0, key=f"f_{col}")
            if selected != "전체":
                filtered = filtered[filtered[col] == selected]

st.sidebar.markdown(f"**필터 결과: {len(filtered)}건**")

# ─────────────────────────────────────────────
# 컬럼 분류
# ─────────────────────────────────────────────
cat_cols = [c for c in filtered.columns
            if (filtered[c].dtype == "object" or str(filtered[c].dtype) == "category")
            and filtered[c].nunique() <= 20]
num_cols = [c for c in filtered.columns if pd.api.types.is_numeric_dtype(filtered[c])]
date_cols = [c for c in filtered.columns
             if pd.api.types.is_datetime64_any_dtype(filtered[c]) and filtered[c].notna().any()]

# ─────────────────────────────────────────────
# KPI 카드
# ─────────────────────────────────────────────
st.markdown("---")

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("테스트 건수", f"{len(filtered)}건")
k2.metric("시제품 수", f"{filtered[join_key].nunique()}개")

if "판정결과" in filtered.columns and len(filtered) > 0:
    pass_n = (filtered["판정결과"] == "적합").sum()
    k3.metric("적합률", f"{pass_n / len(filtered) * 100:.1f}%")
if "pH" in filtered.columns:
    k4.metric("평균 pH", f"{filtered['pH'].mean():.2f}")
if "점도_cP" in filtered.columns:
    k5.metric("평균 점도", f"{filtered['점도_cP'].mean():,.0f} cP")

# ─────────────────────────────────────────────
# 탭 구성
# ─────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 판정결과 분석",
    "🔬 수치 분석",
    "🏆 시제품 성적표",
    "📅 시계열 분석",
    "🔥 크로스 분석",
    "📋 원본 데이터",
])

# ══════════════════════════════════════════════
# 탭 1: 판정결과 분석
# ══════════════════════════════════════════════
with tab1:
    if "판정결과" not in filtered.columns:
        st.info("판정결과 컬럼이 없습니다.")
    else:
        color_map = {"적합": "#2ecc71", "경미변화": "#f39c12", "재검토": "#e74c3c"}

        # 전체 판정 분포
        st.markdown("##### 전체 판정결과 분포")
        c1, c2 = st.columns(2)
        with c1:
            counts = filtered["판정결과"].value_counts().reset_index()
            counts.columns = ["판정결과", "건수"]
            fig = px.bar(counts, x="판정결과", y="건수", color="판정결과",
                         color_discrete_map=color_map, text_auto=True)
            fig.update_layout(showlegend=False, height=350)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = px.pie(filtered, names="판정결과", hole=0.4,
                         color="판정결과", color_discrete_map=color_map)
            fig.update_traces(textinfo="percent+label+value")
            fig.update_layout(height=350)
            st.plotly_chart(fig, use_container_width=True)

        # 그룹별 판정 분포
        st.markdown("##### 그룹별 판정결과")
        group_options = [c for c in ["테스트조건", "시제품코드", "제품유형", "목표피부타입", "담당팀", "주요컨셉"]
                         if c in filtered.columns and filtered[c].notna().any()]
        if group_options:
            group_col = st.selectbox("그룹 기준", group_options, key="judge_group")
            cross = pd.crosstab(filtered[group_col], filtered["판정결과"])
            cross_pct = cross.div(cross.sum(axis=1), axis=0) * 100

            c1, c2 = st.columns(2)
            with c1:
                fig = px.bar(cross.reset_index().melt(id_vars=group_col),
                             x=group_col, y="value", color="판정결과",
                             color_discrete_map=color_map, text_auto=True,
                             title="건수 기준", barmode="group")
                fig.update_layout(yaxis_title="건수", height=400)
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                fig = px.bar(cross_pct.reset_index().melt(id_vars=group_col),
                             x=group_col, y="value", color="판정결과",
                             color_discrete_map=color_map,
                             title="비율 기준 (%)", barmode="stack")
                fig.update_layout(yaxis_title="%", height=400)
                st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════
# 탭 2: 수치 분석
# ══════════════════════════════════════════════
with tab2:
    target_nums = [c for c in ["점도_cP", "pH", "색상변화등급", "보관온도", "보관기간_주"] if c in filtered.columns]

    if not target_nums:
        st.info("수치 컬럼이 없습니다.")
    else:
        st.markdown("##### 기초 통계량")
        st.dataframe(filtered[target_nums].describe().round(2), use_container_width=True)

        st.markdown("##### 판정결과별 수치 비교")
        sel_num = st.selectbox("수치 컬럼", target_nums, key="num_sel")

        c1, c2 = st.columns(2)
        with c1:
            if "판정결과" in filtered.columns:
                fig = px.box(filtered, x="판정결과", y=sel_num, color="판정결과",
                             color_discrete_map={"적합": "#2ecc71", "경미변화": "#f39c12", "재검토": "#e74c3c"},
                             title=f"판정결과별 {sel_num}")
                fig.update_layout(showlegend=False, height=400)
                st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = px.histogram(filtered, x=sel_num, nbins=15, marginal="box",
                               title=f"{sel_num} 분포", color_discrete_sequence=["#3498db"])
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)

        # 수치 컬럼 상관관계
        if len(target_nums) >= 2:
            st.markdown("##### 수치 컬럼 상관관계")
            corr = filtered[target_nums].corr()
            fig = px.imshow(corr, text_auto=".2f", color_continuous_scale="RdBu_r",
                            title="상관관계 히트맵", aspect="auto")
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)

        # 산점도
        if len(target_nums) >= 2:
            st.markdown("##### 산점도")
            sc1, sc2 = st.columns(2)
            with sc1:
                scatter_x = st.selectbox("X축", target_nums, index=0, key="scatter_x")
            with sc2:
                scatter_y = st.selectbox("Y축", target_nums, index=min(1, len(target_nums)-1), key="scatter_y")
            color_by = "판정결과" if "판정결과" in filtered.columns else None
            fig = px.scatter(filtered, x=scatter_x, y=scatter_y, color=color_by,
                             color_discrete_map={"적합": "#2ecc71", "경미변화": "#f39c12", "재검토": "#e74c3c"},
                             hover_data=["시제품코드", "테스트조건"] if "테스트조건" in filtered.columns else None,
                             title=f"{scatter_x} vs {scatter_y}")
            fig.update_layout(height=450)
            st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════
# 탭 3: 시제품 성적표
# ══════════════════════════════════════════════
with tab3:
    if "판정결과" not in filtered.columns:
        st.info("판정결과 컬럼이 없습니다.")
    else:
        st.markdown("##### 시제품별 적합률 랭킹")

        score = filtered.groupby(join_key).agg(
            테스트수=("판정결과", "size"),
            적합=(("판정결과", lambda x: (x == "적합").sum())),
            경미변화=(("판정결과", lambda x: (x == "경미변화").sum())),
            재검토=(("판정결과", lambda x: (x == "재검토").sum())),
        ).reset_index()
        score["적합률(%)"] = (score["적합"] / score["테스트수"] * 100).round(1)
        score = score.sort_values("적합률(%)", ascending=False)

        # 메타 정보 병합 (중복 제거)
        meta_summary = meta.groupby(join_key).first().reset_index()
        meta_cols_to_add = [c for c in ["제품유형", "제형", "개발단계", "목표피부타입", "주요컨셉", "담당팀"]
                            if c in meta_summary.columns]
        if meta_cols_to_add:
            score = score.merge(meta_summary[[join_key] + meta_cols_to_add], on=join_key, how="left")

        # 적합률 바 차트
        fig = px.bar(score, x=join_key, y="적합률(%)",
                     text="적합률(%)",
                     color="적합률(%)",
                     color_continuous_scale=["#e74c3c", "#f39c12", "#2ecc71"],
                     title="시제품별 적합률")
        fig.update_layout(height=400)
        st.plotly_chart(fig, use_container_width=True)

        # 스택 바 차트
        fig2 = go.Figure()
        for label, color in [("적합", "#2ecc71"), ("경미변화", "#f39c12"), ("재검토", "#e74c3c")]:
            fig2.add_trace(go.Bar(
                x=score[join_key], y=score[label],
                name=label, marker_color=color, text=score[label], textposition="inside",
            ))
        fig2.update_layout(barmode="stack", title="시제품별 판정 구성", height=400)
        st.plotly_chart(fig2, use_container_width=True)

        # 성적표 테이블
        st.markdown("##### 상세 성적표")
        st.dataframe(score, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════
# 탭 4: 시계열 분석
# ══════════════════════════════════════════════
with tab4:
    if not date_cols:
        st.info("날짜 컬럼이 없습니다.")
    else:
        date_col = st.selectbox("날짜 컬럼", date_cols, key="ts_col") if len(date_cols) > 1 else date_cols[0]
        valid = filtered[filtered[date_col].notna()].copy()

        if valid.empty:
            st.info("유효한 날짜 데이터가 없습니다.")
        else:
            daily = valid.groupby(valid[date_col].dt.date).size().reset_index()
            daily.columns = ["날짜", "건수"]
            daily["누적"] = daily["건수"].cumsum()

            fig = go.Figure()
            fig.add_trace(go.Bar(x=daily["날짜"], y=daily["건수"], name="일별 건수", marker_color="#3498db"))
            fig.add_trace(go.Scatter(
                x=daily["날짜"], y=daily["누적"], name="누적",
                mode="lines+markers", yaxis="y2", line=dict(color="#e74c3c"),
            ))
            fig.update_layout(
                title=f"일별 테스트 추이 ({date_col} 기준)",
                yaxis=dict(title="건수"),
                yaxis2=dict(title="누적", overlaying="y", side="right"),
                height=400,
            )
            st.plotly_chart(fig, use_container_width=True)

            # 판정결과별 시계열
            if "판정결과" in valid.columns:
                st.markdown("##### 판정결과별 일별 추이")
                daily_judge = valid.groupby([valid[date_col].dt.date, "판정결과"]).size().reset_index()
                daily_judge.columns = ["날짜", "판정결과", "건수"]
                fig2 = px.bar(daily_judge, x="날짜", y="건수", color="판정결과",
                              color_discrete_map={"적합": "#2ecc71", "경미변화": "#f39c12", "재검토": "#e74c3c"},
                              barmode="stack", title="판정결과별 일별 분포")
                fig2.update_layout(height=400)
                st.plotly_chart(fig2, use_container_width=True)

# ══════════════════════════════════════════════
# 탭 5: 크로스 분석
# ══════════════════════════════════════════════
with tab5:
    if len(cat_cols) >= 2:
        cl, cr = st.columns(2)
        with cl:
            x_axis = st.selectbox("X축", cat_cols, index=0, key="cross_x")
        with cr:
            y_axis = st.selectbox("Y축", cat_cols,
                                  index=min(1, len(cat_cols) - 1), key="cross_y")

        cross = pd.crosstab(filtered[y_axis], filtered[x_axis])
        fig = px.imshow(cross, text_auto=True, color_continuous_scale="Blues",
                        title=f"{y_axis} × {x_axis} 교차표", aspect="auto")
        fig.update_layout(height=450)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("크로스 분석을 위해 2개 이상의 카테고리 컬럼이 필요합니다.")

# ══════════════════════════════════════════════
# 탭 6: 원본 데이터
# ══════════════════════════════════════════════
with tab6:
    st.markdown("##### 조인된 전체 데이터")
    st.dataframe(filtered, use_container_width=True, height=400)

    buffer = BytesIO()
    filtered.to_excel(buffer, index=False, engine="openpyxl")
    st.download_button(
        label="📥 필터링 데이터 다운로드 (Excel)",
        data=buffer.getvalue(),
        file_name="analysis_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
