from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd


INPUT_FILE = "请替换为上传后的Excel文件路径.xlsx"
OUTPUT_FILE = "库存分类结果.xlsx"

MONTH_COLUMNS = ["2026-01", "2026-02", "2026-03", "2026-04", "2026-05"]
STOCK_COLUMNS = ["期末库存", "采购在途", "销售未出库数量", "预定订单数量"]
ABCXYZ_SORT_ORDER = ["AX", "BX", "CX", "AY", "BY", "CY", "AZ", "BZ", "CZ"]
HALF_YEAR_DAYS = 180
LONG_SHELF_TURNOVER_DAYS = 45
SHORT_SHELF_TURNOVER_DAYS = 30
LOW_STOCK_COVER_MONTHS = 0.6
HIGH_STOCK_COVER_MONTHS = 2.5
VOLATILITY_WARNING_CV = 0.7
INVENTORY_RISK_LOW_SALES = 1
INVENTORY_RISK_LONG_COVER_MONTHS = 3
SALES_DECLINE_RECENT_RATIO = 0.7
SALES_DECLINE_FULL_RATIO = 0.5

OUTPUT_COLUMNS = [
    "罗诊物料号",
    "物料",
    "物料名称",
    "库存余量",
    "ABCXYZ",
    "效期classification",
    "效期",
    "销售后效期",
    "sales_amount",
    "sales_amount_contribution_pct",
    "mean_sales",
    "库存覆盖月数",
    "CV",
    "采购模式",
    "采购模式判断依据",
    "库存分析结果",
    "库存分析判断依据",
    "销量持续下滑",
    "销售后效期分类",
    "SKU",
    "期末库存",
    "采购在途",
    "销售未出库数量",
    "预定订单数量",
    "ABC",
    "XYZ",
    "2026-01",
    "2026-02",
    "2026-03",
    "2026-04",
    "2026-05",
    "reason",
    "采购单价",
]


def clean_column_name(column: object) -> str:
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
    return clean_column_names(pd.read_excel(input_file))


def add_reason(existing: str, reason: str) -> str:
    if not reason:
        return existing
    if pd.isna(existing) or existing == "":
        return reason
    if reason in str(existing).split("; "):
        return str(existing)
    return f"{existing}; {reason}"


def add_reason_where(df: pd.DataFrame, mask: pd.Series, reason: str) -> None:
    df.loc[mask, "reason"] = df.loc[mask, "reason"].apply(lambda value: add_reason(value, reason))


def require_columns(df: pd.DataFrame, required_columns: Iterable[str]) -> None:
    missing_columns = [column for column in required_columns if column not in df.columns]
    if missing_columns:
        raise ValueError(f"Excel 缺少必要字段: {', '.join(missing_columns)}")


def ensure_material_name(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "物料名称" not in df.columns:
        df["物料名称"] = df["SKU"]
    return df

def analysis_month_start(analysis_date: date) -> date:
    return date(analysis_date.year, analysis_date.month, 1)


def shelf_life_category(days: pd.Series) -> pd.Series:
    values = pd.to_numeric(days, errors="coerce")
    return pd.Series(
        np.select(
            [values < HALF_YEAR_DAYS, values >= HALF_YEAR_DAYS],
            ["短", "长"],
            default="效期缺失",
        ),
        index=days.index,
    )


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


def calculate_inventory_balance(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for column in STOCK_COLUMNS:
        if column not in df.columns:
            df[column] = 0
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0)

    df["库存余量"] = df["期末库存"] + df["采购在途"] - df["销售未出库数量"] - df["预定订单数量"]
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
        [df["cumulative_sales_amount_ratio"] <= 0.80, df["cumulative_sales_amount_ratio"] <= 0.95],
        ["A", "B"],
        default="C",
    )
    return df


def classify_xyz(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["XYZ"] = np.select(
        [df["CV"] < 0.5, (df["CV"] >= 0.5) & (df["CV"] < 1.0)],
        ["X", "Y"],
        default="Z",
    )
    add_reason_where(df, df["CV"].isna(), "无稳定需求")
    return df


def classify_shelf_life(df: pd.DataFrame, analysis_date: date | None = None) -> pd.DataFrame:
    df = df.copy()
    if analysis_date is None:
        analysis_date = date.today()

    if "效期" not in df.columns:
        df["效期"] = np.nan

    df["效期"] = pd.to_numeric(df["效期"], errors="coerce")
    df["效期分类"] = shelf_life_category(df["效期"])
    df["动态当前效期"] = df["效期"]
    df["效期classification"] = shelf_life_category(df["效期"])
    df["默认周转天数"] = np.where(df["效期classification"] == "长", LONG_SHELF_TURNOVER_DAYS, SHORT_SHELF_TURNOVER_DAYS)
    df.loc[df["效期classification"] == "效期缺失", "默认周转天数"] = np.nan
    df["销售后效期"] = df["动态当前效期"] - df["默认周转天数"]
    df["销售后效期分类"] = shelf_life_category(df["销售后效期"])

    add_reason_where(df, df["效期"].isna(), "效期缺失")
    add_reason_where(df, df["效期classification"] == "短", "短效期")
    add_reason_where(df, df["销售后效期分类"] == "短", "销售后短效期")
    return df


def classify_procurement_mode(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["库存覆盖月数"] = df["库存余量"] / df["mean_sales"].replace(0, np.nan)

    short_shelf_life = df["销售后效期分类"] == "短"
    no_sales = df["mean_sales"] <= 0
    enough_stock_cover = df["库存覆盖月数"] > HIGH_STOCK_COVER_MONTHS
    low_stock_cover = df["库存覆盖月数"] < LOW_STOCK_COVER_MONTHS
    within_stock_cover = df["库存覆盖月数"].le(HIGH_STOCK_COVER_MONTHS)
    high_volatility = df["CV"] >= VOLATILITY_WARNING_CV
    stable_enough = df["CV"].lt(VOLATILITY_WARNING_CV).fillna(False)
    regular_candidate = (~short_shelf_life) & (~no_sales) & within_stock_cover & stable_enough

    procurement_conditions = [
        short_shelf_life,
        no_sales,
        enough_stock_cover,
        high_volatility,
        regular_candidate,
    ]
    df["采购模式"] = np.select(
        procurement_conditions,
        ["按单采购", "按单采购", "按单采购", "按单采购", "常规备货"],
        default="按单采购",
    )

    procurement_basis = np.select(
        procurement_conditions,
        [
            "销售处理后短效期",
            "无销量",
            "库存覆盖超过2.5个月",
            "销量波动系数达到0.7，暂不常规备货",
            "库存覆盖不超过2.5个月且销量较稳定",
        ],
        default="数据不足",
    )
    df["采购模式判断依据"] = pd.Series(procurement_basis, index=df.index, dtype="object")
    df.loc[regular_candidate & low_stock_cover, "采购模式判断依据"] = "库存覆盖低于0.6个月且销量较稳定，备货信号强"
    df.loc[regular_candidate & (~low_stock_cover), "采购模式判断依据"] = "库存覆盖不超过2.5个月且销量较稳定"
    df["MTO/MTS"] = df["采购模式"]
    return df


def classify_inventory_risk(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    has_stock = df["库存余量"] > 0
    no_sales = df["mean_sales"] <= 0
    low_sales = (df["mean_sales"] > 0) & (df["mean_sales"] < INVENTORY_RISK_LOW_SALES)
    short_after_sales = df["销售后效期分类"] == "短"
    long_after_sales = df["销售后效期分类"] == "长"
    long_cover = df["库存覆盖月数"] > INVENTORY_RISK_LONG_COVER_MONTHS

    first_month = df[MONTH_COLUMNS[0]]
    third_month = df[MONTH_COLUMNS[2]]
    fourth_month = df[MONTH_COLUMNS[3]]
    last_month = df[MONTH_COLUMNS[4]]
    full_period_decline = (first_month > 0) & (last_month <= first_month * SALES_DECLINE_FULL_RATIO)
    recent_three_month_decline = (
        (third_month > fourth_month)
        & (fourth_month > last_month)
        & (third_month > 0)
        & (last_month <= third_month * SALES_DECLINE_RECENT_RATIO)
    )
    sales_declining = full_period_decline | recent_three_month_decline

    risk_conditions = [
        ~has_stock,
        has_stock & no_sales,
        has_stock & long_cover & sales_declining,
        has_stock & long_cover,
        has_stock & low_sales,
        has_stock & short_after_sales,
        has_stock & long_after_sales,
    ]
    df["库存分析结果"] = np.select(
        risk_conditions,
        [
            "无可售库存",
            "滞销风险-无销量库存",
            "滞销风险-覆盖偏长且销量下滑",
            "滞销风险-库存覆盖偏长",
            "滞销风险-极低销量库存",
            "需要关注-销售后短效期",
            "库存健康",
        ],
        default="需要关注",
    )
    df["库存分析判断依据"] = np.select(
        risk_conditions,
        [
            "库存余量小于等于0",
            "有库存但近五月无销量",
            "库存覆盖月数超过3个月，且销量持续下滑",
            "库存覆盖月数超过3个月",
            "有库存但近五月平均销量低于1",
            "库存覆盖未超过3个月，但销售处理后效期为短",
            "库存覆盖未超过3个月，且销售处理后效期为长",
        ],
        default="未命中明确库存风险规则",
    )
    df["销量持续下滑"] = sales_declining
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


def classify_inventory(
    df: pd.DataFrame,
    analysis_date: date | None = None,
) -> pd.DataFrame:
    df = clean_column_names(df)
    df = ensure_material_name(df)
    df = calculate_mean_sales_and_cv(df)
    df = calculate_inventory_balance(df)
    df = classify_abc(df)
    df = classify_xyz(df)
    df = classify_shelf_life(df, analysis_date=analysis_date)
    df = classify_procurement_mode(df)
    df = classify_inventory_risk(df)
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
    result_df = classify_inventory(read_inventory_excel(input_file))
    write_inventory_result(result_df, output_file)
    print(f"已输出库存分类结果: {Path(output_file).resolve()}")


if __name__ == "__main__":
    main()
