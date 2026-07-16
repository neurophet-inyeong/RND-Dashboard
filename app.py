from __future__ import annotations

from pathlib import Path
from typing import Iterable

import pandas as pd
import plotly.express as px
import streamlit as st


PROJECT_DIR = Path("projects")
DEFAULT_SHEETS = ["개요", "사업비", "마일스톤", "정량지표", "정량지표상세"]


@st.cache_data(show_spinner=False)
def list_excel_files(project_dir: Path) -> list[Path]:
    patterns = ("*.xlsx", "*.xls", "*.xlsm")
    files: list[Path] = []
    for pattern in patterns:
        files.extend(project_dir.glob(pattern))
    return sorted(files)


def normalize_col(col: object) -> str:
    text = str(col)
    text = text.replace("\n", " ").replace("\r", " ").strip()
    text = " ".join(text.split())
    return text


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    renamed = {col: normalize_col(col) for col in df.columns}
    df = df.rename(columns=renamed)

    unnamed_count = 0
    final_cols: list[str] = []
    for col in df.columns:
        if col.lower().startswith("unnamed") or col == "nan":
            unnamed_count += 1
            final_cols.append(f"미지정컬럼_{unnamed_count}")
        else:
            final_cols.append(col)
    df.columns = final_cols
    return df


@st.cache_data(show_spinner=False)
def read_sheet(file_path: str, sheet_name: str, header_row: int = 0) -> pd.DataFrame:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
    df = normalize_columns(df)
    return df


def try_parse_numeric(df: pd.DataFrame) -> pd.DataFrame:
    parsed = df.copy()
    for col in parsed.columns:
        if parsed[col].dtype == object:
            cleaned = (
                parsed[col]
                .astype(str)
                .str.replace(",", "", regex=False)
                .str.replace("%", "", regex=False)
                .str.strip()
            )
            numeric = pd.to_numeric(cleaned, errors="coerce")
            if numeric.notna().sum() >= max(2, int(len(parsed) * 0.4)):
                parsed[col] = numeric
    return parsed


def find_date_candidates(df: pd.DataFrame) -> list[str]:
    date_cols: list[str] = []
    for col in df.columns:
        series = pd.to_datetime(df[col], errors="coerce")
        if series.notna().sum() >= max(2, int(len(df) * 0.4)):
            date_cols.append(col)
    return date_cols


def numeric_columns(df: pd.DataFrame) -> list[str]:
    return [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]


def category_columns(df: pd.DataFrame, exclude: Iterable[str]) -> list[str]:
    exclude_set = set(exclude)
    cols = []
    for col in df.columns:
        if col in exclude_set:
            continue
        if pd.api.types.is_object_dtype(df[col]) or pd.api.types.is_string_dtype(df[col]):
            cols.append(col)
    return cols


def build_project_summary(file_path: Path) -> dict[str, object]:
    summary = {
        "프로젝트": file_path.stem,
        "총사업비": 0.0,
        "달성": 0,
        "미달성": 0,
    }

    try:
        budget = read_sheet(str(file_path), "사업비")
        budget = try_parse_numeric(budget)
        budget_cols = [c for c in budget.columns if "총 사업비" in c or c == "총사업비"]
        if budget_cols:
            summary["총사업비"] = float(budget[budget_cols[0]].fillna(0).sum())
    except Exception:
        pass

    try:
        metrics = read_sheet(str(file_path), "정량지표")
        reached = 0
        missed = 0
        for col in metrics.columns:
            if metrics[col].dtype != object:
                continue
            values = metrics[col].astype(str)
            reached += int(values.str.contains("달성", na=False).sum())
            missed += int(values.str.contains("미달성", na=False).sum())
        summary["달성"] = reached
        summary["미달성"] = missed
    except Exception:
        pass

    return summary


def format_krw(value: float) -> str:
    return f"{value:,.0f} 원"


def render_overview(files: list[Path]) -> None:
    st.subheader("전체 프로젝트 성과 요약")

    rows = [build_project_summary(path) for path in files]
    if not rows:
        st.info("요약할 데이터가 없습니다.")
        return

    summary_df = pd.DataFrame(rows)

    total_budget = summary_df["총사업비"].sum()
    total_reached = int(summary_df["달성"].sum())
    total_missed = int(summary_df["미달성"].sum())

    c1, c2, c3 = st.columns(3)
    c1.metric("총 사업비 합계", format_krw(total_budget))
    c2.metric("달성 건수", f"{total_reached:,}")
    c3.metric("미달성 건수", f"{total_missed:,}")

    budget_chart = px.bar(
        summary_df,
        x="프로젝트",
        y="총사업비",
        title="프로젝트별 총 사업비",
        text_auto=True,
    )
    budget_chart.update_layout(yaxis_title="원")
    st.plotly_chart(budget_chart, use_container_width=True)

    status_df = pd.DataFrame(
        {
            "상태": ["달성", "미달성"],
            "건수": [total_reached, total_missed],
        }
    )
    status_chart = px.pie(status_df, names="상태", values="건수", title="정량지표 달성 현황")
    st.plotly_chart(status_chart, use_container_width=True)

    display_df = summary_df.copy()
    display_df["총사업비"] = display_df["총사업비"].map(format_krw)
    st.dataframe(display_df, use_container_width=True)


def render_sheet_explorer(file_path: Path, sheet_name: str) -> None:
    st.subheader("시트 데이터 탐색")

    preview = pd.read_excel(str(file_path), sheet_name=sheet_name, header=None, nrows=8)
    with st.expander("헤더 자동 인식이 맞지 않으면 헤더 행을 바꿔주세요", expanded=False):
        st.dataframe(preview, use_container_width=True)

    header_row = st.number_input(
        "헤더 행 번호 (0부터 시작)", min_value=0, max_value=7, value=0, step=1
    )

    df = read_sheet(str(file_path), sheet_name, header_row=int(header_row))
    df = df.dropna(how="all")
    if df.empty:
        st.warning("데이터가 비어 있습니다.")
        return

    df = try_parse_numeric(df)
    date_cols = find_date_candidates(df)
    num_cols = numeric_columns(df)

    st.caption(f"행: {len(df):,}개, 열: {len(df.columns):,}개")
    st.dataframe(df.head(200), use_container_width=True)

    if not num_cols:
        st.info("수치형 컬럼이 없어 차트를 그릴 수 없습니다. 다른 시트 또는 헤더 행을 선택해주세요.")
        return

    st.markdown("### 차트 설정")
    x_candidates = date_cols if date_cols else df.columns.tolist()
    default_x = x_candidates[0]

    x_col = st.selectbox("X축", options=x_candidates, index=0)
    y_col = st.selectbox("Y축(수치)", options=num_cols, index=0)

    cat_cols = category_columns(df, exclude=[x_col, y_col])
    group_col = st.selectbox("그룹(선택)", options=["없음", *cat_cols], index=0)

    chart_type = st.radio("차트 유형", options=["라인", "막대", "산점도"], horizontal=True)

    chart_df = df.copy()
    if x_col in date_cols:
        chart_df[x_col] = pd.to_datetime(chart_df[x_col], errors="coerce")

    if chart_type == "라인":
        if group_col != "없음":
            fig = px.line(chart_df, x=x_col, y=y_col, color=group_col, markers=True)
        else:
            fig = px.line(chart_df, x=x_col, y=y_col, markers=True)
    elif chart_type == "막대":
        if group_col != "없음":
            grouped = chart_df.groupby([x_col, group_col], dropna=False, as_index=False)[y_col].sum()
            fig = px.bar(grouped, x=x_col, y=y_col, color=group_col, barmode="group")
        else:
            grouped = chart_df.groupby(x_col, dropna=False, as_index=False)[y_col].sum()
            fig = px.bar(grouped, x=x_col, y=y_col)
    else:
        if group_col != "없음":
            fig = px.scatter(chart_df, x=x_col, y=y_col, color=group_col)
        else:
            fig = px.scatter(chart_df, x=x_col, y=y_col)

    fig.update_layout(margin=dict(l=12, r=12, t=36, b=12), yaxis_title=y_col)
    st.plotly_chart(fig, use_container_width=True)


st.set_page_config(page_title="과제 성과지표 대시보드", layout="wide")
st.title("과제별 성과지표 시각화 앱")
st.caption("projects 폴더의 엑셀 파일을 자동 인식해 성과 데이터를 요약/시각화합니다.")

if not PROJECT_DIR.exists():
    st.error("projects 폴더가 존재하지 않습니다. 워크스페이스 루트에 projects 폴더를 만들어주세요.")
    st.stop()

excel_files = list_excel_files(PROJECT_DIR)
if not excel_files:
    st.warning("projects 폴더에 엑셀 파일(.xlsx/.xls/.xlsm)이 없습니다.")
    st.stop()

render_overview(excel_files)

st.divider()

selected_file = st.selectbox("프로젝트 파일 선택", options=excel_files, format_func=lambda p: p.name)

available_sheets = pd.ExcelFile(str(selected_file)).sheet_names
sheet_options = [s for s in DEFAULT_SHEETS if s in available_sheets]
extra_sheets = [s for s in available_sheets if s not in sheet_options]
sheet_options.extend(extra_sheets)

selected_sheet = st.selectbox("시트 선택", options=sheet_options)
render_sheet_explorer(selected_file, selected_sheet)
