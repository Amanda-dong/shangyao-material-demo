from __future__ import annotations

from pathlib import Path

import pandas as pd


INPUT_FILE = "请替换为上传后的Excel文件路径.xlsx"
OUTPUT_FILE = "/Users/amandadongsmacbookpro/Downloads/销量季节性抽样结果.xlsx"
MONTH_COLUMN = "月份"
MATERIAL_COLUMN = "物料"
QUANTITY_COLUMN = "交货数量"
PURCHASE_ORDER_COLUMN = "采购订单编号"
TARGET_MONTHS = [1, 2, 3, 4, 5]


def clean_column_name(column: object) -> str:
    cleaned = str(column)
    for char in ("\n", "\r", "\t", " ", '"', "'", "“", "”", "‘", "’"):
        cleaned = cleaned.replace(char, "")
    return cleaned


def clean_column_names(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [clean_column_name(column) for column in df.columns]
    return df


def require_columns(df: pd.DataFrame, required_columns: list[str]) -> None:
    missing_columns = [column for column in required_columns if column not in df.columns]
    if missing_columns:
        raise ValueError(f"缺少必要字段: {', '.join(missing_columns)}")


def normalize_month(value: object) -> int | None:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if "-" in text:
        text = text.split("-")[-1]
    text = text.replace("月", "")
    try:
        month = int(float(text))
    except ValueError:
        return None
    return month if 1 <= month <= 12 else None


def prepare_long_sales_df(df: pd.DataFrame) -> tuple[pd.DataFrame, str, str, str, str]:
    df = clean_column_names(df)
    require_columns(df, [MATERIAL_COLUMN, MONTH_COLUMN, QUANTITY_COLUMN, PURCHASE_ORDER_COLUMN])
    material_column = MATERIAL_COLUMN
    order_column = PURCHASE_ORDER_COLUMN
    quantity_column = QUANTITY_COLUMN
    month_column = MONTH_COLUMN

    df[quantity_column] = pd.to_numeric(df[quantity_column], errors="coerce").fillna(0)
    df = df[df[quantity_column] != 0].copy()

    df["月份_标准"] = df[month_column].apply(normalize_month)
    df = df[df["月份_标准"].isin(TARGET_MONTHS)].copy()
    return df, material_column, order_column, quantity_column, month_column


def read_xslx(input_file: str | Path = INPUT_FILE) -> pd.DataFrame:
    """Read long-format Excel and keep first 5 materials and first 5 order IDs."""
    raw_df = pd.read_excel(input_file)
    df, material_column, order_column, _, _ = prepare_long_sales_df(raw_df)

    first_materials = df[material_column].dropna().drop_duplicates().head(5)
    first_orders = df[order_column].dropna().drop_duplicates().head(5)
    sample_df = df[
        df[material_column].isin(first_materials)
        & df[order_column].isin(first_orders)
    ].copy()
    return sample_df


def build_material_monthly_demand(sample_df: pd.DataFrame) -> pd.DataFrame:
    df, material_column, _, quantity_column, _ = prepare_long_sales_df(sample_df)
    monthly = (
        df.groupby([material_column, "月份_标准"], dropna=False)[quantity_column]
        .sum()
        .reset_index(name="月需求数量")
    )
    monthly_pivot = monthly.pivot_table(
        index=material_column,
        columns="月份_标准",
        values="月需求数量",
        fill_value=0,
        aggfunc="sum",
    ).reset_index()
    monthly_pivot.columns = [material_column, *[f"{int(column)}月需求" for column in monthly_pivot.columns[1:]]]

    month_columns = [column for column in monthly_pivot.columns if column.endswith("月需求")]
    monthly_pivot["1-5月总需求"] = monthly_pivot[month_columns].sum(axis=1)
    monthly_pivot["1-5月平均需求"] = monthly_pivot[month_columns].mean(axis=1)
    monthly_pivot["最大月份"] = monthly_pivot[month_columns].idxmax(axis=1)
    monthly_pivot["最大/平均"] = monthly_pivot[month_columns].max(axis=1) / monthly_pivot["1-5月平均需求"].replace(0, pd.NA)
    monthly_pivot["初步季节性提示"] = monthly_pivot["最大/平均"].apply(
        lambda value: "可能有季节性波动" if pd.notna(value) and value >= 1.5 else "暂未看到明显季节性"
    )
    return monthly_pivot


def build_order_contribution(sample_df: pd.DataFrame) -> pd.DataFrame:
    df, material_column, order_column, quantity_column, _ = prepare_long_sales_df(sample_df)
    contribution = (
        df.groupby([material_column, "月份_标准", order_column], dropna=False)[quantity_column]
        .sum()
        .reset_index(name="订单需求数量")
    )
    month_total = contribution.groupby([material_column, "月份_标准"], dropna=False)["订单需求数量"].transform("sum")
    contribution["订单贡献占比"] = (contribution["订单需求数量"] / month_total.replace(0, pd.NA) * 100).round(1)
    contribution = contribution.sort_values([material_column, "月份_标准", "订单需求数量"], ascending=[True, True, False])
    return contribution


def main() -> None:
    input_file = INPUT_FILE
    output_file = OUTPUT_FILE
    if not Path(input_file).exists():
        raise FileNotFoundError(f"请先设置有效输入文件路径: {input_file}")

    sample_df = read_xslx(input_file)
    contribution_df = build_order_contribution(sample_df)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        contribution_df.to_excel(writer, index=False, sheet_name="订单贡献", float_format="%.1f")

    print(f"已输出: {Path(output_file).resolve()}")


if __name__ == "__main__":
    main()
