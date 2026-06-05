from __future__ import annotations

from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd


INPUT_FILE = "请替换为上传后的Excel文件路径.xlsx"
OUTPUT_FILE = "库存分类结果.xlsx"

MONTH_COLUMNS = ["2026-01", "2026-02", "2026-03", "2026-04", "2026-05"]
OUTPUT_COLUMNS = [
    "SKU",
    "期末库存",
    "采购在途",
    "销售未出库数量",
    "预定订单数量",
    "平均销售",
    "ABCXYZ",
    "ABC",
    "XYZ",
    "效期",
    "效期classification",
    "2026-01",
    "2026-02",
    "2026-03",
    "2026-04",
    "2026-05",
    "sales_amount_contribution_pct",
    "mean_sales",
    "CV",
    "sales_amount",
    "MTO/MTS",
    "reason",
    "采购单价",
]

ABCXYZ_SORT_ORDER = ["AX", "BX", "CX", "AY", "BY", "CY", "AZ", "BZ", "CZ"]


def clean_column_name(column: object) -> str:
    """Normalize Excel headers such as newlines, spaces, and quote marks."""
    cleaned = str(column)
    for char in ("\n", "\r", "\t", " ", '"', "'", "“", "”", "‘", "’"):
        cleaned = cleaned.replace(char, "")
    return cleaned


def clean_column_names(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [clean_column_name(column) for column in df.columns]

    if "SKU" not in df.columns and len(df.columns) > 0 and df.columns[0].startswith("Unnamed"):
        df = df.rename(columns={df.columns[0]: "SKU"})

    return df


def read_inventory_excel(input_file: str | Path) -> pd.DataFrame:
    df = pd.read_excel(input_file)
    return clean_column_names(df)


def add_reason(existing: str, reason: str) -> str:
    if not reason:
        return existing
    if pd.isna(existing) or existing == "":
        return reason
    return f"{existing}; {reason}"


def add_reason_where(df: pd.DataFrame, mask: pd.Series, reason: str) -> None:
    df.loc[mask, "reason"] = df.loc[mask, "reason"].apply(lambda value: add_reason(value, reason))


def require_columns(df: pd.DataFrame, required_columns: Iterable[str]) -> None:
    missing_columns = [column for column in required_columns if column not in df.columns]
    if missing_columns:
        raise ValueError(f"Excel 缺少必要字段: {', '.join(missing_columns)}")


def calculate_mean_sales_and_cv(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    require_columns(df, ["SKU", "采购单价", *MONTH_COLUMNS])

    for column in [*MONTH_COLUMNS, "采购单价"]:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0)

    df["reason"] = ""
    df["mean_sales"] = df[MONTH_COLUMNS].mean(axis=1)
    df["CV"] = df[MONTH_COLUMNS].std(axis=1) / df["mean_sales"]
    df.loc[df["mean_sales"] == 0, "CV"] = np.nan
    add_reason_where(df, df["mean_sales"] == 0, "近5个月无销量")

    return df


def classify_abc(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["sales_amount"] = df["mean_sales"] * df["采购单价"]
    df = df.sort_values("sales_amount", ascending=False, kind="mergesort").reset_index(drop=True)

    total_sales_amount = df["sales_amount"].sum()
    if total_sales_amount > 0:
        df["sales_amount_contribution_pct"] = df["sales_amount"] / total_sales_amount * 100
        df["cumulative_sales_amount_ratio"] = df["sales_amount"].cumsum() / total_sales_amount
    else:
        df["sales_amount_contribution_pct"] = 0.0
        df["cumulative_sales_amount_ratio"] = 1.0

    df["ABC"] = np.select(
        [
            df["cumulative_sales_amount_ratio"] <= 0.80,
            df["cumulative_sales_amount_ratio"] <= 0.95,
        ],
        ["A", "B"],
        default="C",
    )
    return df


def classify_xyz(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["XYZ"] = np.select(
        [
            df["CV"] < 0.5,
            (df["CV"] >= 0.5) & (df["CV"] < 1.0),
        ],
        ["X", "Y"],
        default="Z",
    )
    add_reason_where(df, df["CV"].isna(), "无稳定需求")
    return df


def classify_stock_type(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    zero_mask = (df["mean_sales"] == 0) | (df["mean_sales"] < 1) | (df["CV"] > 2)
    df["stock_type"] = np.where(zero_mask, "0", "非0")
    df["MTO/MTS"] = np.where(zero_mask, "MTO", "MTS")

    add_reason_where(df, df["mean_sales"] == 0, "无销量")
    add_reason_where(df, (df["mean_sales"] > 0) & (df["mean_sales"] < 1), "低需求")
    add_reason_where(df, df["CV"] > 2, "高波动")

    return df


def classify_shelf_life(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    if "效期" not in df.columns:
        df["效期"] = np.nan

    df["效期"] = pd.to_numeric(df["效期"], errors="coerce")
    df["效期classification"] = np.select(
        [
            df["效期"] < 90,
            (df["效期"] >= 90) & (df["效期"] <= 180),
            (df["效期"] >= 181) & (df["效期"] <= 365),
            df["效期"] > 365,
        ],
        ["高危<90", "预警90-180", "健康181-365", "安全>365"],
        default="效期缺失",
    )

    add_reason_where(df, df["效期"].isna(), "效期缺失")
    add_reason_where(df, df["效期"] < 90, "效期高危")

    return df


def sort_for_purchase_review(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["ABCXYZ"] = df["ABC"] + df["XYZ"]
    sort_rank = {label: index for index, label in enumerate(ABCXYZ_SORT_ORDER)}
    df["_sort_rank"] = df["ABCXYZ"].map(sort_rank).fillna(len(sort_rank))
    return (
        df.sort_values(["_sort_rank", "sales_amount"], ascending=[True, False], kind="mergesort")
        .drop(columns=["_sort_rank"])
        .reset_index(drop=True)
    )


def classify_inventory(df: pd.DataFrame) -> pd.DataFrame:
    df = clean_column_names(df)
    df = calculate_mean_sales_and_cv(df)
    df = classify_abc(df)
    df = classify_xyz(df)
    df = classify_shelf_life(df)
    df = classify_stock_type(df)
    df = sort_for_purchase_review(df)
    return df


def build_output_df(df: pd.DataFrame) -> pd.DataFrame:
    output_df = df.copy()
    for column in OUTPUT_COLUMNS:
        if column not in output_df.columns:
            output_df[column] = np.nan
    return output_df[OUTPUT_COLUMNS]


def write_inventory_result(df: pd.DataFrame, output_file: str | Path) -> None:
    build_output_df(df).to_excel(output_file, index=False)


def main() -> None:
    input_file = INPUT_FILE
    output_file = OUTPUT_FILE

    if not Path(input_file).exists():
        raise FileNotFoundError(f"请先设置有效的输入文件路径: {input_file}")

    raw_df = read_inventory_excel(input_file)
    result_df = classify_inventory(raw_df)
    write_inventory_result(result_df, output_file)
    print(f"已输出库存分类结果: {Path(output_file).resolve()}")


if __name__ == "__main__":
    main()
