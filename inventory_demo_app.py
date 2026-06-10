from __future__ import annotations

from io import BytesIO

import altair as alt
import pandas as pd
import streamlit as st

from inventory_classification_demo import build_output_df, classify_inventory


SHELF_LIFE_ORDER = ["短", "长", "效期缺失"]
PROCUREMENT_ORDER = ["常规备货", "按单采购"]
XYZ_ORDER = ["X", "Y", "Z"]
ABCXYZ_ORDER = ["AX", "BX", "CX", "AY", "BY", "CY", "AZ", "BZ", "CZ"]

DISPLAY_COLUMN_NAMES = {
    "罗诊物料号": "罗诊物料号",
    "物料": "物料",
    "SKU": "物料编码",
    "物料名称": "物料名称",
    "ABCXYZ": "物料分类",
    "sales_amount_contribution_pct": "销售额贡献占比",
    "mean_sales": "近五月平均销量",
    "sales_amount": "销售额",
    "CV": "销量波动系数",
    "库存覆盖月数": "库存覆盖月数",
    "效期classification": "效期分类",
    "效期": "效期",
    "效期分类": "原始效期分类",
    "动态当前效期": "动态当前效期",
    "销售后效期": "销售处理后效期",
    "销售后效期分类": "销售后效期分类",
    "库存余量": "库存余量",
    "库存分析结果": "库存分析",
    "库存分析判断依据": "库存分析判断依据",
    "销量持续下滑": "销量持续下滑",
    "采购模式": "采购模式",
    "采购模式判断依据": "采购模式判断依据",
    "reason": "原因",
}

NUMERIC_COLUMN_FORMATS = {
    "sales_amount_contribution_pct": "%.2f%%",
    "mean_sales": "%.2f",
    "CV": "%.2f",
    "库存覆盖月数": "%.2f",
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
    "库存覆盖月数": "0.00",
    "sales_amount": "0.00",
    "采购单价": "0.00",
    "库存余量": "0.00",
    "效期": "0",
    "动态当前效期": "0",
    "销售后效期": "0",
}

TABLE_BASE_COLUMNS = [
    "罗诊物料号",
    "物料",
    "物料名称",
    "库存余量",
    "ABCXYZ",
    "效期",
    "效期classification",
    "销售后效期",
    "销售后效期分类",
    "sales_amount",
    "sales_amount_contribution_pct",
    "mean_sales",
    "库存覆盖月数",
    "CV",
    "采购模式",
    "采购模式判断依据",
    "库存分析结果",
    "库存分析判断依据",
]

CLASSIFICATION_COLUMNS = TABLE_BASE_COLUMNS
SHELF_LIFE_TABLE_COLUMNS = TABLE_BASE_COLUMNS
PROCUREMENT_COLUMNS = [
    "罗诊物料号",
    "物料",
    "物料名称",
    "库存余量",
    "ABCXYZ",
    "效期",
    "效期classification",
    "销售后效期",
    "销售后效期分类",
    "sales_amount",
    "sales_amount_contribution_pct",
    "mean_sales",
    "库存覆盖月数",
    "CV",
    "采购模式",
    "采购模式判断依据",
]
INVENTORY_RISK_COLUMNS = [
    "罗诊物料号",
    "物料",
    "物料名称",
    "库存余量",
    "ABCXYZ",
    "效期",
    "效期classification",
    "销售后效期",
    "销售后效期分类",
    "sales_amount",
    "sales_amount_contribution_pct",
    "mean_sales",
    "库存覆盖月数",
    "CV",
    "销量持续下滑",
    "库存分析结果",
    "库存分析判断依据",
]

MAIN_TABLE_COLUMNS = [
    "罗诊物料号",
    "物料",
    "SKU",
    "物料名称",
    "库存余量",
    "效期",
    "效期classification",
    "销售后效期",
    "销售后效期分类",
    "mean_sales",
    "库存覆盖月数",
    "CV",
    "销量持续下滑",
    "ABCXYZ",
    "采购模式",
    "采购模式判断依据",
    "库存分析结果",
    "库存分析判断依据",
    "sales_amount",
    "sales_amount_contribution_pct",
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


def render_upload_panel() -> object:
    st.markdown('<div class="section-label">上传文件</div>', unsafe_allow_html=True)
    return st.file_uploader("上传库存/销量数据", type=["xlsx", "xls"], key="inventory_file")


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


def simplified_inventory_status(series: pd.Series) -> pd.Series:
    values = series.astype(str)
    labels = values.map(
        lambda value: "滞销风险"
        if value.startswith("滞销风险")
        else "需要关注"
        if value.startswith("需要关注")
        else "无可售库存"
        if value == "无可售库存"
        else "库存健康"
        if value == "库存健康"
        else "其他"
    )
    return pd.Series(
        pd.Categorical(
            labels,
            categories=["库存健康", "需要关注", "滞销风险", "无可售库存", "其他"],
            ordered=True,
        ),
        index=series.index,
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


def dedupe_columns(columns: list[str]) -> list[str]:
    seen = set()
    result = []
    for column in columns:
        if column in seen:
            continue
        seen.add(column)
        result.append(column)
    return result


def render_table(df: pd.DataFrame) -> None:
    df = df.loc[:, dedupe_columns(list(df.columns))]
    st.dataframe(df, use_container_width=True, hide_index=True, column_config=dataframe_column_config(df))


def sort_xyz_table(df: pd.DataFrame) -> pd.DataFrame:
    table_df = df.copy()
    table_df["_xyz_rank"] = table_df["XYZ"].map({label: index for index, label in enumerate(XYZ_ORDER)}).fillna(len(XYZ_ORDER))
    return table_df.sort_values(["_xyz_rank", "ABCXYZ", "sales_amount"], ascending=[True, True, False]).drop(columns=["_xyz_rank"])


def shelf_life_groups(df: pd.DataFrame) -> list[tuple[str, pd.DataFrame]]:
    labels = [label for label in SHELF_LIFE_ORDER if label in set(df["效期classification"])]
    return [(label, sort_xyz_table(df[df["效期classification"] == label])) for label in labels]


def filter_options(df: pd.DataFrame, column: str) -> list[str]:
    if column not in df.columns:
        return []
    return sorted([str(value) for value in df[column].dropna().unique()])


def apply_multi_filter(df: pd.DataFrame, column: str, selected: list[str]) -> pd.DataFrame:
    if not selected or column not in df.columns:
        return df
    return df[df[column].astype(str).isin(selected)]


def render_main_decision_table(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-label">物料判断总表</div>', unsafe_allow_html=True)
    st.markdown(
        '<p class="group-caption">一张表合并查看物料分类、采购模式和库存分析。先用筛选器定位问题物料，再看判断依据。</p>',
        unsafe_allow_html=True,
    )
    filter_columns = st.columns(4)
    with filter_columns[0]:
        procurement_filter = st.multiselect("采购模式", filter_options(df, "采购模式"))
    with filter_columns[1]:
        inventory_filter = st.multiselect("库存分析", filter_options(df, "库存分析结果"))
    with filter_columns[2]:
        shelf_filter = st.multiselect("效期分类", filter_options(df, "效期classification"))
    with filter_columns[3]:
        material_filter = st.multiselect("物料分类", filter_options(df, "ABCXYZ"))

    filtered_df = df.copy()
    filtered_df = apply_multi_filter(filtered_df, "采购模式", procurement_filter)
    filtered_df = apply_multi_filter(filtered_df, "库存分析结果", inventory_filter)
    filtered_df = apply_multi_filter(filtered_df, "效期classification", shelf_filter)
    filtered_df = apply_multi_filter(filtered_df, "ABCXYZ", material_filter)

    st.markdown(f'<p class="group-caption">当前显示 {len(filtered_df)} / {len(df)} 个物料</p>', unsafe_allow_html=True)
    render_table(filtered_df[[column for column in MAIN_TABLE_COLUMNS if column in filtered_df.columns]])


def render_rule_summary() -> None:
    st.markdown('<div class="section-label">规则说明</div>', unsafe_allow_html=True)
    st.markdown("\n        <p class=\"group-caption\"><b>关键公式</b><br>\n        近五月平均销量 = 2026-01 至 2026-05 的平均销量。销量波动系数 = 近五个月销量标准差 / 近五月平均销量。库存覆盖月数 = 库存余量 / 近五月平均销量。销售额 = 近五月平均销量 x 采购单价。\n        </p>\n\n        <p class=\"group-caption\"><b>ABCXYZ 物料分类</b><br>\n        ABC 看销售额贡献：按销售额从高到低排序，累计贡献 &lt;=80% 为 A，80%-95% 为 B，95% 以后为 C。XYZ 看销量稳定性：CV &lt;0.5 为 X，0.5-1.0 为 Y，CV &gt;=1.0 或无稳定需求为 Z。ABCXYZ 只是价值和稳定性的标签，不直接决定采购模式。\n        </p>\n\n        <p class=\"group-caption\"><b>采购模式判断</b><br>\n        采购模式用于判断平时是否进入常规备货池。销售处理后效期为短、近五月无销量、库存覆盖月数 &gt;2.5、或 CV &gt;=0.7 时，判为按单采购；库存覆盖月数 &lt;=2.5 且 CV &lt;0.7 时，判为常规备货。库存覆盖月数 &lt;0.6 会被标记为备货信号强。\n        </p>\n\n        <p class=\"group-caption\"><b>库存分析</b><br>\n        库存分析独立于采购模式，重点看现有库存是否可能形成滞销。有库存但近五月无销量，判为滞销风险-无销量库存。库存覆盖月数 &gt;3，判为滞销风险-库存覆盖偏长；如果同时销量持续下滑，判为滞销风险-覆盖偏长且销量下滑。库存覆盖月数 &lt;=3 且销售处理后效期为长，判为库存健康；库存覆盖月数 &lt;=3 但销售处理后效期为短，判为需要关注-销售后短效期。\n        </p>\n\n        <p class=\"group-caption\"><b>销量持续下滑定义</b><br>\n        满足任一条件即视为销量持续下滑：五月销量 &lt;= 一月销量的 50%；或三月 &gt; 四月 &gt; 五月，且五月销量 &lt;= 三月销量的 70%。\n        </p>\n", unsafe_allow_html=True)


def render_classification_table(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-label">分类结果</div>', unsafe_allow_html=True)
    st.markdown(
        '<p class="group-caption">稳定性：销量波动系数 = 近 5 个月销量标准差 / 近 5 个月平均销量。价值：销售额 = 平均销量 x 采购单价。库存覆盖月数 = 库存余量 / 近五月平均销量。采购模式先看销售处理后效期和无销量，再看库存覆盖月数和销量波动系数：销售处理后短效期、库存覆盖超过 2.5 个月或 CV 达到 0.7 时，暂不进入常规备货。</p>',
        unsafe_allow_html=True,
    )
    render_table(df[[column for column in CLASSIFICATION_COLUMNS if column in df.columns]])


def render_shelf_life_tables(df: pd.DataFrame) -> None:
    groups = shelf_life_groups(df)
    if not groups:
        return
    st.markdown('<div class="section-label">按效期分组的物料</div>', unsafe_allow_html=True)
    st.markdown(
        '<p class="group-caption">效期分类标准：短 <180 天；长 >=180 天。动态当前效期按上传表中的效期判断。</p>',
        unsafe_allow_html=True,
    )
    tabs = st.tabs([f"{label} ({len(group_df)})" for label, group_df in groups])
    for tab, (_, group_df) in zip(tabs, groups):
        with tab:
            render_table(group_df[[column for column in SHELF_LIFE_TABLE_COLUMNS if column in group_df.columns]])


def render_inventory_risk_table(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-label">库存风险分析</div>', unsafe_allow_html=True)
    st.markdown(
        '<p class="group-caption">库存分析独立于采购模式：库存覆盖月数 >3 个月视为覆盖偏长；若同时销量持续下滑，则判为覆盖偏长且销量下滑。销量持续下滑定义为：五月销量 <= 一月销量的 50%，或三月 > 四月 > 五月且五月 <= 三月的 70%。销售处理后短效期作为关注因素。</p>',
        unsafe_allow_html=True,
    )
    render_table(df[[column for column in INVENTORY_RISK_COLUMNS if column in df.columns]])


def render_procurement_tables(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-label">采购模式判断</div>', unsafe_allow_html=True)
    st.markdown('<p class="group-caption">采购模式规则：销售处理后效期为短按单采购；无销量按单采购；库存覆盖月数 >2.5 按单采购；CV >=0.7 按单采购；库存覆盖月数 <=2.5 且 CV <0.7 进入常规备货。库存覆盖 <0.6 是强备货信号。</p>', unsafe_allow_html=True)
    metric_columns = st.columns(2)
    metric_columns[0].metric("常规备货", int((df["采购模式"] == "常规备货").sum()))
    metric_columns[1].metric("按单采购", int((df["采购模式"] == "按单采购").sum()))

    if "采购模式判断依据" in df.columns:
        render_distribution("采购判断依据分布", df["采购模式判断依据"])

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

    inventory_file = render_upload_panel()
    if inventory_file is None:
        render_empty_state()
        return

    try:
        raw_df = pd.read_excel(inventory_file)
        result_df = classify_inventory(raw_df)
        output_df = build_output_df(result_df)

        inventory_attention = output_df["库存分析结果"].astype(str).str.startswith(("滞销风险", "需要关注")).sum()
        metric_columns = st.columns(4)
        metric_columns[0].metric("物料总数", len(output_df))
        metric_columns[1].metric("常规备货", int((output_df["采购模式"] == "常规备货").sum()))
        metric_columns[2].metric("按单采购", int((output_df["采购模式"] == "按单采购").sum()))
        metric_columns[3].metric("库存需关注", int(inventory_attention))

        chart_columns = st.columns(2, gap="large")
        with chart_columns[0]:
            render_distribution("采购模式分布", output_df["采购模式"], PROCUREMENT_ORDER)
        with chart_columns[1]:
            render_distribution("库存分析分布", simplified_inventory_status(output_df["库存分析结果"]))

        render_main_decision_table(output_df)

        with st.expander("ABCXYZ 分类明细", expanded=False):
            render_distribution("物料分类分布", output_df["ABCXYZ"], ABCXYZ_ORDER)
            render_classification_table(output_df)

        with st.expander("效期分组明细", expanded=False):
            render_shelf_life_tables(output_df)

        with st.expander("采购模式分表", expanded=False):
            render_procurement_tables(output_df)

        with st.expander("库存分析明细", expanded=False):
            render_inventory_risk_table(output_df)

        with st.expander("规则说明", expanded=False):
            render_rule_summary()

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
