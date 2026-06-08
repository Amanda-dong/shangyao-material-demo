from __future__ import annotations

from datetime import date
from io import BytesIO

import altair as alt
import pandas as pd
import streamlit as st

from inventory_classification_demo import build_output_df, classify_inventory


SHELF_LIFE_ORDER = ["短", "长", "效期缺失"]
PROCUREMENT_ORDER = ["常规备货", "按单采购"]
XYZ_ORDER = ["X", "Y", "Z"]

DISPLAY_COLUMN_NAMES = {
    "SKU": "物料编码",
    "物料名称": "物料名称",
    "ABCXYZ": "物料分类",
    "sales_amount_contribution_pct": "销售价值贡献占比",
    "mean_sales": "平均销量",
    "sales_amount": "销售价值",
    "CV": "销量波动系数",
    "效期classification": "动态当前效期分类",
    "效期": "效期",
    "效期分类": "原始效期分类",
    "动态当前效期": "动态当前效期",
    "销售后效期": "销售处理后效期",
    "销售后效期分类": "销售后效期分类",
    "库存余量": "库存余量",
    "库存分析结果": "库存分析结果",
    "采购模式": "采购模式",
    "reason": "原因",
}

NUMERIC_COLUMN_FORMATS = {
    "sales_amount_contribution_pct": "%.2f%%",
    "mean_sales": "%.2f",
    "CV": "%.2f",
    "sales_amount": "%.2f",
    "采购单价": "%.2f",
    "库存余量": "%.2f",
    "效期": "%.0f",
    "动态当前效期": "%.0f",
    "销售后效期": "%.0f",
}

EXCEL_NUMBER_FORMATS = {
    "sales_amount_contribution_pct": '0.00"%"',
    "mean_sales": "0.00",
    "CV": "0.00",
    "sales_amount": "0.00",
    "采购单价": "0.00",
    "库存余量": "0.00",
    "效期": "0",
    "动态当前效期": "0",
    "销售后效期": "0",
}

CLASSIFICATION_COLUMNS = [
    "物料名称",
    "SKU",
    "ABCXYZ",
    "效期classification",
    "效期",
    "动态当前效期",
    "销售后效期",
    "sales_amount_contribution_pct",
    "mean_sales",
    "CV",
    "sales_amount",
    "库存余量",
    "采购模式",
    "库存分析结果",
]

SHELF_LIFE_TABLE_COLUMNS = [
    "物料名称",
    "SKU",
    "ABCXYZ",
    "效期classification",
    "效期",
    "效期分类",
    "动态当前效期",
    "销售后效期",
    "销售后效期分类",
    "库存余量",
    "库存分析结果",
    "采购模式",
    "mean_sales",
    "reason",
]

INVENTORY_RISK_COLUMNS = [
    "物料名称",
    "SKU",
    "库存余量",
    "效期",
    "动态当前效期",
    "销售后效期",
    "销售后效期分类",
    "mean_sales",
    "库存分析结果",
    "reason",
]

PROCUREMENT_COLUMNS = [
    "物料名称",
    "SKU",
    "ABCXYZ",
    "效期classification",
    "销售后效期分类",
    "库存余量",
    "mean_sales",
    "采购模式",
    "库存分析结果",
    "reason",
]


def inject_styles() -> None:
    st.markdown(
        """
        <style>
        :root {
            --ink: #f4f7fb;
            --muted: #9aa7b8;
            --line: #263244;
            --panel: #111827;
            --panel-2: #0b1220;
            --accent: #33c3d6;
            --accent-soft: #0d3440;
        }
        .stApp {
            background: radial-gradient(circle at 18% 0%, rgba(51, 195, 214, 0.14), transparent 34%),
                        linear-gradient(180deg, #05070b 0%, #0b1018 44%, #070a0f 100%);
            color: var(--ink);
        }
        [data-testid="stHeader"] { background: transparent; }
        .block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 1280px; }
        .hero, .panel, .upload-row {
            border: 1px solid var(--line);
            background: linear-gradient(135deg, rgba(17, 24, 39, 0.98) 0%, rgba(11, 18, 32, 0.98) 100%);
            border-radius: 8px;
            box-shadow: 0 14px 34px rgba(0, 0, 0, 0.26);
        }
        .hero { padding: 26px 30px; margin-bottom: 20px; border-color: rgba(51, 195, 214, 0.24); }
        .hero h1 { margin: 0 0 8px; font-size: 34px; line-height: 1.15; color: var(--ink); }
        .hero p, .panel-copy, .group-caption, .file-caption { color: var(--muted); font-size: 13px; }
        .upload-row { padding: 16px 18px; margin-bottom: 18px; }
        .panel { padding: 18px; min-height: 132px; }
        .panel-title { margin: 0 0 6px; color: var(--ink); font-size: 17px; font-weight: 700; }
        .panel-copy { margin: 0 0 14px; }
        .status-pill { display: inline-block; border: 1px solid rgba(51, 195, 214, 0.35); background: var(--accent-soft); color: var(--accent); border-radius: 999px; padding: 4px 10px; font-size: 12px; font-weight: 700; margin-top: 6px; }
        [data-testid="stMetric"] { background: linear-gradient(180deg, #111827 0%, #0c1320 100%); border: 1px solid var(--line); border-radius: 8px; padding: 14px 16px; }
        [data-testid="stMetricValue"] { color: var(--accent); font-weight: 800; }
        [data-testid="stFileUploader"] { background: var(--panel-2); border: 1px dashed rgba(51, 195, 214, 0.42); border-radius: 8px; padding: 12px; }
        [data-testid="stFileUploader"] * { color: var(--ink); }
        .stDownloadButton button, .stButton button { border-radius: 6px; border: 1px solid var(--accent); background: var(--accent); color: #031017; font-weight: 700; }
        .section-label { margin: 26px 0 10px; font-size: 18px; font-weight: 800; color: var(--ink); }
        .group-caption { margin: 0 0 10px; }
        [data-testid="stDataFrame"] { border: 1px solid var(--line); border-radius: 8px; overflow: hidden; }
        .stTabs [data-baseweb="tab-list"] { gap: 8px; }
        .stTabs [data-baseweb="tab"] { background: #111827; border: 1px solid var(--line); border-radius: 6px; color: var(--muted); padding: 8px 14px; }
        .stTabs [aria-selected="true"] { color: var(--accent); border-color: rgba(51, 195, 214, 0.55); background: var(--accent-soft); }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header() -> None:
    st.markdown(
        """
        <div class="hero">
            <h1>物料分类 Demo</h1>
            <p>上传库存数据后，系统会计算 ABCXYZ、动态效期、库存余量、采购模式和库存风险。</p>
            <span class="status-pill">仅通过上传文件读取数据</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_upload_panel() -> tuple[object, date]:
    st.markdown('<div class="section-label">上传与分析日期</div>', unsafe_allow_html=True)
    columns = st.columns([1.6, 0.8], gap="large")
    with columns[0]:
        inventory_file = st.file_uploader("上传库存/销量数据", type=["xlsx", "xls"], key="inventory_file")
    with columns[1]:
        analysis_date = st.date_input("分析日期", value=date.today())
    return inventory_file, analysis_date


def render_empty_state() -> None:
    st.markdown(
        """
        <div class="panel">
            <p class="panel-title">等待上传</p>
            <p class="panel-copy">请先上传库存/销量数据。系统只会读取你在页面上传的这一张表。</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_distribution(title: str, series: pd.Series, order: list[str] | None = None) -> None:
    counts = series.value_counts().rename_axis("分类").reset_index(name="数量")
    if order:
        counts["_rank"] = counts["分类"].map({label: index for index, label in enumerate(order)}).fillna(len(order))
        counts = counts.sort_values("_rank").drop(columns=["_rank"])
    st.markdown(f'<div class="section-label">{title}</div>', unsafe_allow_html=True)
    summary = " / ".join(f"{row['分类']}: {int(row['数量'])}" for _, row in counts.iterrows())
    st.markdown(f'<p class="group-caption">{summary}</p>', unsafe_allow_html=True)
    chart = (
        alt.Chart(counts)
        .mark_bar(color="#33c3d6", cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
        .encode(
            x=alt.X("分类:N", sort=None, axis=alt.Axis(labelAngle=0, title=None)),
            y=alt.Y("数量:Q", axis=alt.Axis(title=None, grid=True)),
            tooltip=["分类:N", "数量:Q"],
        )
        .properties(height=260)
    )
    st.altair_chart(chart, use_container_width=True)


def dataframe_column_config(df: pd.DataFrame) -> dict[str, object]:
    config = {}
    for column in df.columns:
        label = DISPLAY_COLUMN_NAMES.get(column, column)
        if column in NUMERIC_COLUMN_FORMATS:
            config[column] = st.column_config.NumberColumn(label, format=NUMERIC_COLUMN_FORMATS[column])
        else:
            config[column] = st.column_config.TextColumn(label)
    return config


def render_table(df: pd.DataFrame) -> None:
    st.dataframe(df, use_container_width=True, hide_index=True, column_config=dataframe_column_config(df))


def sort_xyz_table(df: pd.DataFrame) -> pd.DataFrame:
    table_df = df.copy()
    table_df["_xyz_rank"] = table_df["XYZ"].map({label: index for index, label in enumerate(XYZ_ORDER)}).fillna(len(XYZ_ORDER))
    return table_df.sort_values(["_xyz_rank", "ABCXYZ", "sales_amount"], ascending=[True, True, False]).drop(columns=["_xyz_rank"])


def shelf_life_groups(df: pd.DataFrame) -> list[tuple[str, pd.DataFrame]]:
    labels = [label for label in SHELF_LIFE_ORDER if label in set(df["效期classification"])]
    return [(label, sort_xyz_table(df[df["效期classification"] == label])) for label in labels]


def render_classification_table(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-label">分类结果</div>', unsafe_allow_html=True)
    st.markdown(
        '<p class="group-caption">稳定性：销量波动系数 = 近 5 个月销量标准差 / 近 5 个月平均销量。价值：销售价值 = 平均销量 x 采购单价。效期按分析日期动态扣减本月 1 号至当天的天数。</p>',
        unsafe_allow_html=True,
    )
    render_table(df[[column for column in CLASSIFICATION_COLUMNS if column in df.columns]])


def render_shelf_life_tables(df: pd.DataFrame) -> None:
    groups = shelf_life_groups(df)
    if not groups:
        return
    st.markdown('<div class="section-label">按效期分组的物料</div>', unsafe_allow_html=True)
    st.markdown(
        '<p class="group-caption">效期分类标准：短 <180 天；长 >=180 天。效期、动态当前效期、销售处理后效期都会显示。</p>',
        unsafe_allow_html=True,
    )
    tabs = st.tabs([f"{label} ({len(group_df)})" for label, group_df in groups])
    for tab, (_, group_df) in zip(tabs, groups):
        with tab:
            render_table(group_df[[column for column in SHELF_LIFE_TABLE_COLUMNS if column in group_df.columns]])


def render_inventory_risk_table(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-label">库存风险分析</div>', unsafe_allow_html=True)
    st.markdown(
        '<p class="group-caption">滞销风险会同时看库存余量、销售处理后效期和销量。销售处理后仍为长效期且有销量时通常不算滞销；销售处理后为短效期、低销量或无销量库存会被标记为风险。</p>',
        unsafe_allow_html=True,
    )
    render_table(df[[column for column in INVENTORY_RISK_COLUMNS if column in df.columns]])


def render_procurement_tables(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-label">采购模式判断</div>', unsafe_allow_html=True)
    st.markdown('<p class="group-caption">常规备货 = 可进入备货池；按单采购 = 常规不备货，有需求再看单处理。</p>', unsafe_allow_html=True)
    metric_columns = st.columns(2)
    metric_columns[0].metric("常规备货", int((df["采购模式"] == "常规备货").sum()))
    metric_columns[1].metric("按单采购", int((df["采购模式"] == "按单采购").sum()))

    for mode in PROCUREMENT_ORDER:
        st.markdown(f'<div class="section-label">{mode}物料</div>', unsafe_allow_html=True)
        mode_df = df[df["采购模式"] == mode]
        render_table(mode_df[[column for column in PROCUREMENT_COLUMNS if column in mode_df.columns]])


def get_excel_column_letter(column_index: int) -> str:
    letters = ""
    while column_index:
        column_index, remainder = divmod(column_index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def apply_excel_number_formats(writer: pd.ExcelWriter, df: pd.DataFrame, sheet_name: str) -> None:
    worksheet = writer.sheets[sheet_name]
    for column_name, number_format in EXCEL_NUMBER_FORMATS.items():
        if column_name not in df.columns:
            continue
        column_index = df.columns.get_loc(column_name) + 1
        column_letter = get_excel_column_letter(column_index)
        for cell in worksheet[column_letter][1:]:
            cell.number_format = number_format


def write_sheet(writer: pd.ExcelWriter, df: pd.DataFrame, sheet_name: str) -> None:
    df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    apply_excel_number_formats(writer, df, sheet_name[:31])


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        write_sheet(writer, df, "总表")
        for label, group_df in shelf_life_groups(df):
            write_sheet(writer, group_df, label)
        risk_df = df[[column for column in INVENTORY_RISK_COLUMNS if column in df.columns]]
        write_sheet(writer, risk_df, "库存风险分析")
    return output.getvalue()


def run_demo() -> None:
    st.set_page_config(page_title="物料分类 Demo", layout="wide")
    inject_styles()
    render_header()

    inventory_file, analysis_date = render_upload_panel()
    if inventory_file is None:
        render_empty_state()
        return

    try:
        raw_df = pd.read_excel(inventory_file)
        result_df = classify_inventory(raw_df, analysis_date=analysis_date)
        output_df = build_output_df(result_df)

        metric_columns = st.columns(4)
        metric_columns[0].metric("物料数量", len(output_df))
        metric_columns[1].metric("常规备货", int((output_df["采购模式"] == "常规备货").sum()))
        metric_columns[2].metric("按单采购", int((output_df["采购模式"] == "按单采购").sum()))
        metric_columns[3].metric("滞销风险", int(output_df["库存分析结果"].astype(str).str.startswith("滞销风险").sum()))

        chart_columns = st.columns(2, gap="large")
        with chart_columns[0]:
            render_distribution("效期分层", output_df["效期classification"], SHELF_LIFE_ORDER)
        with chart_columns[1]:
            render_distribution("库存风险分布", output_df["库存分析结果"])

        render_classification_table(output_df)
        render_shelf_life_tables(output_df)
        render_procurement_tables(output_df)
        render_inventory_risk_table(output_df)

        st.download_button(
            "下载分类结果 Excel（含效期与库存风险表）",
            data=to_excel_bytes(output_df),
            file_name="分类结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:
        st.error(str(exc))


if __name__ == "__main__":
    run_demo()
