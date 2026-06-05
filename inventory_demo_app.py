from __future__ import annotations

from io import BytesIO

import altair as alt
import pandas as pd
import streamlit as st
from inventory_classification_demo import build_output_df, classify_inventory


SHELF_LIFE_ORDER = ["高危<90", "预警90-180", "健康181-365", "安全>365", "效期缺失"]
SHELF_LIFE_DISPLAY_NAMES = {
    "高危<90": "高危",
    "预警90-180": "预警",
    "健康181-365": "健康",
    "安全>365": "安全",
    "效期缺失": "缺失",
}
XYZ_ORDER = ["X", "Y", "Z"]
DISPLAY_COLUMN_NAMES = {
    "SKU": "物料名称",
    "ABCXYZ": "物料分类",
    "mean_sales": "平均销量",
    "sales_amount_contribution_pct": "销售价值贡献占比",
    "sales_amount": "销售价值",
    "效期classification": "效期分类",
}
SHELF_LIFE_TABLE_COLUMNS = [
    "SKU",
    "ABCXYZ",
    "效期classification",
    "ABC",
    "XYZ",
    "MTO/MTS",
    "效期",
    "sales_amount_contribution_pct",
    "mean_sales",
    "CV",
    "sales_amount",
    "reason",
    "采购单价",
]
NUMERIC_COLUMN_FORMATS = {
    "sales_amount_contribution_pct": "%.2f%%",
    "mean_sales": "%.2f",
    "CV": "%.2f",
    "sales_amount": "%.2f",
    "采购单价": "%.2f",
}
EXCEL_NUMBER_FORMATS = {
    "sales_amount_contribution_pct": '0.00"%"',
    "mean_sales": "0.00",
    "CV": "0.00",
    "sales_amount": "0.00",
    "采购单价": "0.00",
}


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
            --soft: #151f2e;
            --accent: #33c3d6;
            --accent-soft: #0d3440;
            --danger: #fb7185;
        }

        .stApp {
            background:
                radial-gradient(circle at 18% 0%, rgba(51, 195, 214, 0.14), transparent 34%),
                linear-gradient(180deg, #05070b 0%, #0b1018 44%, #070a0f 100%);
            color: var(--ink);
        }

        [data-testid="stHeader"] {
            background: transparent;
        }

        .block-container {
            padding-top: 2rem;
            padding-bottom: 3rem;
            max-width: 1280px;
        }

        .hero {
            border: 1px solid rgba(51, 195, 214, 0.24);
            background: linear-gradient(135deg, rgba(17, 24, 39, 0.98) 0%, rgba(11, 18, 32, 0.98) 100%);
            border-radius: 8px;
            padding: 26px 30px;
            box-shadow: 0 18px 52px rgba(0, 0, 0, 0.34);
            margin-bottom: 20px;
        }

        .hero h1 {
            margin: 0 0 8px;
            font-size: 34px;
            line-height: 1.15;
            letter-spacing: 0;
            color: var(--ink);
        }

        .hero p {
            margin: 0;
            color: var(--muted);
            font-size: 15px;
        }

        .panel {
            border: 1px solid var(--line);
            background: var(--panel);
            border-radius: 8px;
            padding: 18px;
            box-shadow: 0 14px 34px rgba(0, 0, 0, 0.26);
            min-height: 132px;
        }

        .upload-row {
            border: 1px solid var(--line);
            background: var(--panel);
            border-radius: 8px;
            padding: 16px 18px;
            box-shadow: 0 14px 34px rgba(0, 0, 0, 0.26);
            margin-bottom: 18px;
        }

        .panel-title {
            margin: 0 0 6px;
            color: var(--ink);
            font-size: 17px;
            font-weight: 700;
        }

        .panel-copy {
            margin: 0 0 14px;
            color: var(--muted);
            font-size: 13px;
        }

        .status-pill {
            display: inline-block;
            border: 1px solid rgba(51, 195, 214, 0.35);
            background: var(--accent-soft);
            color: var(--accent);
            border-radius: 999px;
            padding: 4px 10px;
            font-size: 12px;
            font-weight: 700;
            margin-top: 6px;
        }

        [data-testid="stMetric"] {
            background: linear-gradient(180deg, #111827 0%, #0c1320 100%);
            border: 1px solid var(--line);
            border-radius: 8px;
            padding: 14px 16px;
            box-shadow: 0 12px 28px rgba(0, 0, 0, 0.24);
        }

        [data-testid="stMetricLabel"] p {
            color: var(--muted);
            font-size: 12px;
        }

        [data-testid="stMetricValue"] {
            color: var(--accent);
            font-weight: 800;
        }

        [data-testid="stFileUploader"] {
            background: var(--panel-2);
            border: 1px dashed rgba(51, 195, 214, 0.42);
            border-radius: 8px;
            padding: 12px;
        }

        [data-testid="stFileUploader"] * {
            color: var(--ink);
        }

        [data-testid="stAlert"] {
            background: var(--panel);
            border: 1px solid var(--line);
            color: var(--ink);
        }

        .stDownloadButton button,
        .stButton button {
            border-radius: 6px;
            border: 1px solid var(--accent);
            background: var(--accent);
            color: #031017;
            font-weight: 700;
        }

        .stDownloadButton button:hover,
        .stButton button:hover {
            border-color: #75e4ef;
            background: #75e4ef;
            color: #031017;
        }

        .section-label {
            margin: 26px 0 10px;
            font-size: 18px;
            font-weight: 800;
            color: var(--ink);
        }

        .file-caption {
            color: var(--muted);
            font-size: 13px;
            margin-top: 8px;
        }

        .group-caption {
            color: var(--muted);
            font-size: 13px;
            margin: 0 0 10px;
        }

        [data-testid="stDataFrame"] {
            border: 1px solid var(--line);
            border-radius: 8px;
            overflow: hidden;
        }

        .stTabs [data-baseweb="tab-list"] {
            gap: 8px;
        }

        .stTabs [data-baseweb="tab"] {
            background: #111827;
            border: 1px solid var(--line);
            border-radius: 6px;
            color: var(--muted);
            padding: 8px 14px;
        }

        .stTabs [aria-selected="true"] {
            color: var(--accent);
            border-color: rgba(51, 195, 214, 0.55);
            background: var(--accent-soft);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header() -> None:
    st.markdown(
        """
        <div class="hero">
            <h1>物料分类 Demo</h1>
            <p>上传 Excel 后自动清洗字段、计算销量稳定性、价值贡献、效期分层，并输出采购查看顺序。</p>
            <span class="status-pill">仅通过上传文件读取数据</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_upload_panel() -> object:
    upload_columns = st.columns([1.0, 2.6], gap="large")
    with upload_columns[0]:
        st.markdown(
            """
            <div class="upload-row">
                <p class="panel-title">上传数据</p>
                <p class="panel-copy">支持 .xlsx / .xls。文件上传后才会触发读取和分类。</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with upload_columns[1]:
        return st.file_uploader("上传 Excel 文件", type=["xlsx", "xls"], label_visibility="collapsed")


def render_empty_state() -> None:
    st.markdown(
        """
        <div class="panel">
            <p class="panel-title">等待上传</p>
            <p class="panel-copy">上传后这里会显示物料数量、MTO/MTS、效期风险，以及分类分布图。</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def display_df(df: pd.DataFrame) -> pd.DataFrame:
    display = df.copy()
    if "效期classification" in display.columns:
        display["效期classification"] = display["效期classification"].replace(SHELF_LIFE_DISPLAY_NAMES)
    return display


def render_distribution(title: str, series: pd.Series) -> None:
    series = series.replace(SHELF_LIFE_DISPLAY_NAMES)
    counts = series.value_counts().rename_axis("分类").reset_index(name="数量")
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
    config = {
        column: st.column_config.NumberColumn(
            DISPLAY_COLUMN_NAMES.get(column, column),
            format=number_format,
        )
        for column, number_format in NUMERIC_COLUMN_FORMATS.items()
        if column in df.columns
    }
    for column, label in DISPLAY_COLUMN_NAMES.items():
        if column in df.columns and column not in config:
            config[column] = st.column_config.TextColumn(label)
    return config


def render_table(df: pd.DataFrame) -> None:
    st.dataframe(
        display_df(df),
        use_container_width=True,
        hide_index=True,
        column_config=dataframe_column_config(df),
    )


def sort_xyz_table(df: pd.DataFrame) -> pd.DataFrame:
    xyz_rank = {label: index for index, label in enumerate(XYZ_ORDER)}
    table_df = df.copy()
    table_df["_xyz_rank"] = table_df["XYZ"].map(xyz_rank).fillna(len(xyz_rank))
    table_df = table_df.sort_values(
        ["_xyz_rank", "ABCXYZ", "sales_amount"],
        ascending=[True, True, False],
        kind="mergesort",
    )
    return table_df.drop(columns=["_xyz_rank"])


def shelf_life_groups(df: pd.DataFrame) -> list[tuple[str, pd.DataFrame]]:
    ordered_labels = [label for label in SHELF_LIFE_ORDER if label in set(df["效期classification"])]
    extra_labels = [
        label
        for label in sorted(df["效期classification"].dropna().unique())
        if label not in SHELF_LIFE_ORDER
    ]
    return [
        (label, sort_xyz_table(df[df["效期classification"] == label]))
        for label in [*ordered_labels, *extra_labels]
    ]


def render_shelf_life_tables(df: pd.DataFrame) -> None:
    groups = shelf_life_groups(df)
    if not groups:
        return

    st.markdown('<div class="section-label">按效期分组的物料</div>', unsafe_allow_html=True)
    st.markdown(
        """
        <p class="group-caption">效期分类标准：高危 <90 天；预警 90-180 天；健康 181-365 天；安全 >365 天。90 天以上视为有效期可用范围。</p>
        """,
        unsafe_allow_html=True,
    )
    tabs = st.tabs([f"{SHELF_LIFE_DISPLAY_NAMES.get(label, label)} ({len(group_df)})" for label, group_df in groups])
    for tab, (label, group_df) in zip(tabs, groups):
        with tab:
            xyz_summary = group_df["XYZ"].value_counts().reindex(XYZ_ORDER, fill_value=0)
            st.markdown(
                (
                    f'<p class="group-caption">{SHELF_LIFE_DISPLAY_NAMES.get(label, label)}: '
                    f'X={int(xyz_summary["X"])} / Y={int(xyz_summary["Y"])} / '
                    f'Z={int(xyz_summary["Z"])}</p>'
                ),
                unsafe_allow_html=True,
            )
            columns = [column for column in SHELF_LIFE_TABLE_COLUMNS if column in group_df.columns]
            render_table(group_df[columns])


def render_abcxyz_table(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-label">分类结果</div>', unsafe_allow_html=True)
    st.markdown(
        (
            '<p class="group-caption">稳定性：CV = 近 5 个月销量标准差 / 近 5 个月平均销量，CV 越低需求越稳定。'
            '价值：销售价值 = 平均销量 x 采购单价，用于 ABC 累计贡献度排序。</p>'
        ),
        unsafe_allow_html=True,
    )
    columns = [
        "SKU",
        "ABCXYZ",
        "效期classification",
        "ABC",
        "XYZ",
        "sales_amount_contribution_pct",
        "mean_sales",
        "CV",
        "sales_amount",
        "采购单价",
        "效期",
    ]
    render_table(df[columns])


def render_mto_mts_table(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-label">MTO/MTS 判断</div>', unsafe_allow_html=True)
    st.markdown(
        (
            '<p class="group-caption">在 ABCXYZ 和效期 XYZ 表之后，再看是否常规备货：'
            'MTO = 按单采购，MTS = 常规备货。下面分开显示两张表。</p>'
        ),
        unsafe_allow_html=True,
    )
    metric_columns = st.columns(2)
    metric_columns[0].metric("MTS", int((df["MTO/MTS"] == "MTS").sum()))
    metric_columns[1].metric("MTO", int((df["MTO/MTS"] == "MTO").sum()))

    columns = [
        "SKU",
        "ABCXYZ",
        "效期classification",
        "sales_amount_contribution_pct",
        "mean_sales",
        "CV",
        "MTO/MTS",
        "reason",
    ]
    mts_df = df[df["MTO/MTS"] == "MTS"][columns]
    mto_df = df[df["MTO/MTS"] == "MTO"][columns]

    st.markdown('<div class="section-label">MTS 物料</div>', unsafe_allow_html=True)
    render_table(mts_df)

    st.markdown('<div class="section-label">MTO 物料</div>', unsafe_allow_html=True)
    render_table(mto_df)


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


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        display_df(df).to_excel(writer, index=False, sheet_name="总表")
        apply_excel_number_formats(writer, df, "总表")
        for label, group_df in shelf_life_groups(df):
            safe_label = label.replace("/", "-").replace("\\", "-")[:31]
            display_df(group_df).to_excel(writer, index=False, sheet_name=safe_label)
            apply_excel_number_formats(writer, group_df, safe_label)
    return output.getvalue()


def run_demo() -> None:
    st.set_page_config(page_title="物料分类 Demo", layout="wide")
    inject_styles()
    render_header()

    uploaded_file = render_upload_panel()

    try:
        if uploaded_file is not None:
            raw_df = pd.read_excel(uploaded_file)
            source_name = uploaded_file.name
        else:
            render_empty_state()
            return

        result_df = classify_inventory(raw_df)
        output_df = build_output_df(result_df)

        metric_columns = st.columns(4)
        metric_columns[0].metric("物料数量", len(output_df))
        metric_columns[1].metric("MTS物料", int((output_df["MTO/MTS"] == "MTS").sum()))
        metric_columns[2].metric("MTO物料", int((output_df["MTO/MTS"] == "MTO").sum()))
        metric_columns[3].metric(
            "高危效期物料",
            int((output_df["效期classification"] == "高危<90").sum()),
        )

        chart_columns = st.columns(2, gap="large")
        with chart_columns[0]:
            render_distribution("ABCXYZ 分布", output_df["ABCXYZ"])
        with chart_columns[1]:
            render_distribution("效期分层", output_df["效期classification"])

        render_abcxyz_table(output_df)
        render_shelf_life_tables(output_df)
        render_mto_mts_table(output_df)
        st.download_button(
            "下载分类结果 Excel（含效期分组表）",
            data=to_excel_bytes(output_df),
            file_name="库存分类结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:
        st.error(str(exc))


if __name__ == "__main__":
    run_demo()
