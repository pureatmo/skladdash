from __future__ import annotations

import json
import re
import sqlite3
import uuid
from dataclasses import asdict, dataclass, field
from datetime import date, datetime, timezone
from io import BytesIO
from pathlib import Path
from typing import Any, BinaryIO

import numpy as np
import pandas as pd


KEY_COLUMNS = ["market", "seedbank", "product", "pack"]

LOCAL_SAMPLE_FILES = {
    "usa_sales": Path(r"C:\Users\Admin\Downloads\продажи USA с 01.01.2026 по 14.07.2026.xlsx"),
    "ww_sales": Path(r"C:\Users\Admin\Downloads\продажи WW с 01.01.2026 по 14.07.2026.xlsx"),
    "usa_stock": Path(r"C:\Users\Admin\Downloads\Остатки us.xlsx"),
    "ww_stock": Path(r"C:\Users\Admin\Downloads\Остатки ww.xlsx"),
}

SALES_COLUMNS = {
    "Сидбанк": "seedbank",
    "Продукт": "product",
    "Упаковка": "pack",
    "Количество": "sold_qty",
    "Закупочная стоимость": "purchase_cost",
    "Стоимость товаров": "revenue",
    "Маржа": "margin",
    "Маржинальность": "margin_pct_source",
    "Доля дохода": "income_share_source",
}

STOCK_COLUMNS = {
    "Seedbank": "seedbank",
    "Product": "product",
    "Pack": "pack",
    "Available Qnt.": "current_stock",
}

ABC_METRICS = {
    "revenue": "revenue",
    "margin": "margin",
    "quantity": "sold_qty",
}

NUMERIC_COLUMNS = [
    "sold_qty",
    "purchase_cost",
    "revenue",
    "margin",
    "margin_pct_source",
    "income_share_source",
    "current_stock",
    "period_daily_sales",
    "history_daily_sales",
    "forecast_daily_sales",
    "forecast_period_qty",
    "purchase_horizon_days",
    "horizon_demand_qty",
    "stock_cover_days",
    "target_cover_days",
    "safety_stock_qty",
    "target_stock",
    "reorder_point",
    "raw_order_qty",
    "order_qty",
    "unit_purchase_cost",
    "order_cost_estimate",
    "category_metric",
    "category_share",
    "category_cum_share",
]


@dataclass
class AnalysisConfig:
    period_start: date
    period_end: date
    abc_metric: str = "revenue"
    threshold_a: float = 0.80
    threshold_b: float = 0.95
    threshold_c: float = 0.99
    cover_days_by_category: dict[str, int] = field(
        default_factory=lambda: {"A": 120, "B": 90, "C": 60, "D": 30}
    )
    lead_time_days: int = 45
    safety_stock_days: int = 15
    purchase_horizon_days: int = 90
    min_order_qty: int = 1
    order_multiple: int = 1
    history_weight: float = 0.35
    history_stat: str = "p75"
    include_history_only: bool = True
    only_below_reorder_point: bool = False

    @property
    def period_days(self) -> int:
        return max((self.period_end - self.period_start).days + 1, 1)

    def to_json(self) -> str:
        payload = asdict(self)
        payload["period_start"] = self.period_start.isoformat()
        payload["period_end"] = self.period_end.isoformat()
        return json.dumps(payload, ensure_ascii=False)


def open_connection(db_path: str | Path) -> sqlite3.Connection:
    path = Path(db_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    init_db(conn)
    return conn


def init_db(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS analysis_runs (
            run_uuid TEXT PRIMARY KEY,
            created_at TEXT NOT NULL,
            name TEXT NOT NULL,
            period_start TEXT NOT NULL,
            period_end TEXT NOT NULL,
            period_days INTEGER NOT NULL,
            config_json TEXT NOT NULL,
            notes TEXT,
            items_count INTEGER NOT NULL,
            recommended_items_count INTEGER NOT NULL,
            total_sales_qty REAL NOT NULL,
            total_revenue REAL NOT NULL,
            total_margin REAL NOT NULL,
            total_order_qty REAL NOT NULL,
            total_order_cost_estimate REAL NOT NULL
        );

        CREATE TABLE IF NOT EXISTS analysis_items (
            run_uuid TEXT NOT NULL,
            created_at TEXT NOT NULL,
            market TEXT NOT NULL,
            seedbank TEXT NOT NULL,
            product TEXT NOT NULL,
            pack TEXT NOT NULL,
            sku TEXT NOT NULL,
            sold_qty REAL NOT NULL,
            purchase_cost REAL NOT NULL,
            revenue REAL NOT NULL,
            margin REAL NOT NULL,
            current_stock REAL NOT NULL,
            period_daily_sales REAL NOT NULL,
            history_daily_sales REAL NOT NULL,
            forecast_daily_sales REAL NOT NULL,
            forecast_period_qty REAL NOT NULL,
            unit_purchase_cost REAL NOT NULL,
            abc_category TEXT NOT NULL,
            category_metric REAL NOT NULL,
            category_share REAL NOT NULL,
            category_cum_share REAL NOT NULL,
            target_cover_days REAL NOT NULL,
            lead_time_days REAL NOT NULL,
            safety_stock_days REAL NOT NULL,
            stock_cover_days REAL,
            target_stock REAL NOT NULL,
            reorder_point REAL NOT NULL,
            raw_order_qty REAL NOT NULL,
            order_qty REAL NOT NULL,
            order_cost_estimate REAL NOT NULL,
            stock_status TEXT NOT NULL,
            data_source TEXT NOT NULL,
            history_runs REAL NOT NULL,
            history_last_category TEXT,
            history_last_seen_at TEXT,
            FOREIGN KEY (run_uuid) REFERENCES analysis_runs(run_uuid)
        );

        CREATE INDEX IF NOT EXISTS idx_analysis_items_sku
        ON analysis_items(market, seedbank, product, pack);

        CREATE INDEX IF NOT EXISTS idx_analysis_items_run
        ON analysis_items(run_uuid);
        """
    )
    _ensure_column(conn, "analysis_items", "purchase_horizon_days", "REAL NOT NULL DEFAULT 90")
    _ensure_column(conn, "analysis_items", "horizon_demand_qty", "REAL NOT NULL DEFAULT 0")
    _ensure_column(conn, "analysis_items", "safety_stock_qty", "REAL NOT NULL DEFAULT 0")
    conn.commit()


def _ensure_column(conn: sqlite3.Connection, table_name: str, column_name: str, definition: str) -> None:
    columns = {
        row[1]
        for row in conn.execute(f"PRAGMA table_info({table_name})").fetchall()
    }
    if column_name not in columns:
        conn.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {definition}")


def _read_excel(source: str | Path | BinaryIO, preferred_sheet: str | None = None) -> pd.DataFrame:
    if hasattr(source, "seek"):
        source.seek(0)
    workbook = pd.ExcelFile(source)
    sheet_name = preferred_sheet if preferred_sheet in workbook.sheet_names else workbook.sheet_names[0]
    return pd.read_excel(workbook, sheet_name=sheet_name)


def _clean_text(series: pd.Series) -> pd.Series:
    return (
        series.fillna("")
        .astype(str)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )


def _validate_columns(df: pd.DataFrame, expected: dict[str, str], source_label: str) -> None:
    missing = [column for column in expected if column not in df.columns]
    if missing:
        raise ValueError(f"{source_label}: missing required columns: {', '.join(missing)}")


def _parse_pack_size(pack: str) -> int | None:
    numbers = [int(value) for value in re.findall(r"\d+", str(pack))]
    if not numbers:
        return None
    if "+" in str(pack):
        return sum(numbers)
    return numbers[0]


def read_sales(source: str | Path | BinaryIO, market: str) -> pd.DataFrame:
    raw = _read_excel(source, preferred_sheet="Export")
    _validate_columns(raw, SALES_COLUMNS, f"{market} sales")
    df = raw.rename(columns=SALES_COLUMNS)[list(SALES_COLUMNS.values())].copy()

    for column in ["seedbank", "product", "pack"]:
        df[column] = _clean_text(df[column])
    for column in [
        "sold_qty",
        "purchase_cost",
        "revenue",
        "margin",
        "margin_pct_source",
        "income_share_source",
    ]:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0)

    df = df[df["seedbank"].ne("") & df["product"].ne("") & df["pack"].ne("")]
    df["market"] = market.upper()
    grouped = (
        df.groupby(KEY_COLUMNS, as_index=False)
        .agg(
            sold_qty=("sold_qty", "sum"),
            purchase_cost=("purchase_cost", "sum"),
            revenue=("revenue", "sum"),
            margin=("margin", "sum"),
            income_share_source=("income_share_source", "sum"),
        )
        .sort_values(KEY_COLUMNS)
        .reset_index(drop=True)
    )
    grouped["margin_pct_source"] = np.where(
        grouped["revenue"].gt(0), grouped["margin"] / grouped["revenue"], 0
    )
    grouped["pack_size"] = grouped["pack"].map(_parse_pack_size)
    return grouped


def read_stock(source: str | Path | BinaryIO, market: str) -> pd.DataFrame:
    raw = _read_excel(source)
    _validate_columns(raw, STOCK_COLUMNS, f"{market} stock")
    df = raw.rename(columns=STOCK_COLUMNS)[list(STOCK_COLUMNS.values())].copy()

    for column in ["seedbank", "product", "pack"]:
        df[column] = _clean_text(df[column])
    df["current_stock"] = pd.to_numeric(df["current_stock"], errors="coerce").fillna(0)
    df = df[df["seedbank"].ne("") & df["product"].ne("") & df["pack"].ne("")]
    df["market"] = market.upper()
    return (
        df.groupby(KEY_COLUMNS, as_index=False)
        .agg(current_stock=("current_stock", "sum"))
        .sort_values(KEY_COLUMNS)
        .reset_index(drop=True)
    )


def load_current_data(
    usa_sales: str | Path | BinaryIO,
    ww_sales: str | Path | BinaryIO,
    usa_stock: str | Path | BinaryIO,
    ww_stock: str | Path | BinaryIO,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    sales = pd.concat(
        [read_sales(usa_sales, "USA"), read_sales(ww_sales, "WW")],
        ignore_index=True,
    )
    stock = pd.concat(
        [read_stock(usa_stock, "USA"), read_stock(ww_stock, "WW")],
        ignore_index=True,
    )
    return sales, stock


def get_history_summary(conn: sqlite3.Connection) -> pd.DataFrame:
    init_db(conn)
    rows = pd.read_sql_query(
        """
        SELECT
            run_uuid,
            created_at,
            market,
            seedbank,
            product,
            pack,
            sold_qty,
            revenue,
            margin,
            period_daily_sales,
            forecast_daily_sales,
            unit_purchase_cost,
            order_qty,
            abc_category,
            current_stock
        FROM analysis_items
        """,
        conn,
    )
    if rows.empty:
        return _empty_history()

    rows["created_at"] = pd.to_datetime(rows["created_at"], errors="coerce")
    observed = rows[rows["sold_qty"].gt(0)].copy()

    base = (
        rows.groupby(KEY_COLUMNS, as_index=False)
        .agg(
            history_runs=("run_uuid", "nunique"),
            history_order_qty_avg=("order_qty", "mean"),
            history_last_seen_at=("created_at", "max"),
        )
    )

    if observed.empty:
        observed_agg = _empty_history()[KEY_COLUMNS + [
            "history_daily_sales_avg",
            "history_daily_sales_p75",
            "history_revenue_avg",
            "history_margin_avg",
            "history_unit_purchase_cost",
        ]]
    else:
        observed_agg = (
            observed.groupby(KEY_COLUMNS, as_index=False)
            .agg(
                history_daily_sales_avg=("period_daily_sales", "mean"),
                history_daily_sales_p75=("period_daily_sales", lambda value: value.quantile(0.75)),
                history_revenue_avg=("revenue", "mean"),
                history_margin_avg=("margin", "mean"),
                history_unit_purchase_cost=("unit_purchase_cost", "median"),
            )
        )

    last = (
        rows.sort_values("created_at")
        .groupby(KEY_COLUMNS, as_index=False)
        .tail(1)[KEY_COLUMNS + ["abc_category", "current_stock", "sold_qty"]]
        .rename(
            columns={
                "abc_category": "history_last_category",
                "current_stock": "history_last_stock",
                "sold_qty": "history_last_sold_qty",
            }
        )
    )

    summary = base.merge(observed_agg, on=KEY_COLUMNS, how="left").merge(last, on=KEY_COLUMNS, how="left")
    for column in [
        "history_runs",
        "history_order_qty_avg",
        "history_daily_sales_avg",
        "history_daily_sales_p75",
        "history_revenue_avg",
        "history_margin_avg",
        "history_unit_purchase_cost",
        "history_last_stock",
        "history_last_sold_qty",
    ]:
        summary[column] = pd.to_numeric(summary[column], errors="coerce").fillna(0)
    summary["history_last_seen_at"] = summary["history_last_seen_at"].astype(str).replace("NaT", "")
    summary["history_last_category"] = summary["history_last_category"].fillna("")
    return summary


def _empty_history() -> pd.DataFrame:
    return pd.DataFrame(
        columns=[
            *KEY_COLUMNS,
            "history_runs",
            "history_order_qty_avg",
            "history_last_seen_at",
            "history_daily_sales_avg",
            "history_daily_sales_p75",
            "history_revenue_avg",
            "history_margin_avg",
            "history_unit_purchase_cost",
            "history_last_category",
            "history_last_stock",
            "history_last_sold_qty",
        ]
    )


def compute_analysis(
    sales: pd.DataFrame,
    stock: pd.DataFrame,
    history: pd.DataFrame,
    config: AnalysisConfig,
) -> pd.DataFrame:
    metric_column = ABC_METRICS.get(config.abc_metric, "revenue")

    current_keys = pd.concat(
        [
            sales[KEY_COLUMNS].assign(has_sales_row=True),
            stock[KEY_COLUMNS].assign(has_stock_row=True),
        ],
        ignore_index=True,
    )
    if config.include_history_only and not history.empty:
        current_keys = pd.concat(
            [current_keys, history[KEY_COLUMNS].assign(has_history_row=True)],
            ignore_index=True,
        )
    for flag_column in ["has_sales_row", "has_stock_row", "has_history_row"]:
        if flag_column not in current_keys.columns:
            current_keys[flag_column] = False

    keys = current_keys[KEY_COLUMNS].drop_duplicates().reset_index(drop=True)
    flags = (
        current_keys.groupby(KEY_COLUMNS, as_index=False)
        .agg(
            has_sales_row=("has_sales_row", "max"),
            has_stock_row=("has_stock_row", "max"),
            has_history_row=("has_history_row", "max"),
        )
        .fillna(False)
    )

    df = (
        keys.merge(flags, on=KEY_COLUMNS, how="left")
        .merge(sales, on=KEY_COLUMNS, how="left")
        .merge(stock, on=KEY_COLUMNS, how="left")
        .merge(history, on=KEY_COLUMNS, how="left")
    )

    for column in [
        "sold_qty",
        "purchase_cost",
        "revenue",
        "margin",
        "income_share_source",
        "margin_pct_source",
        "current_stock",
        "history_runs",
        "history_order_qty_avg",
        "history_daily_sales_avg",
        "history_daily_sales_p75",
        "history_revenue_avg",
        "history_margin_avg",
        "history_unit_purchase_cost",
        "history_last_stock",
        "history_last_sold_qty",
    ]:
        if column not in df.columns:
            df[column] = 0
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0)

    for column in ["history_last_category", "history_last_seen_at"]:
        if column not in df.columns:
            df[column] = ""
        df[column] = df[column].fillna("").astype(str)

    df["pack_size"] = df["pack"].map(_parse_pack_size)
    df["sku"] = df[["market", "seedbank", "product", "pack"]].agg(" | ".join, axis=1)
    df["period_daily_sales"] = df["sold_qty"] / config.period_days

    history_daily_column = (
        "history_daily_sales_p75" if config.history_stat == "p75" else "history_daily_sales_avg"
    )
    df["history_daily_sales"] = df[history_daily_column]
    weight = min(max(float(config.history_weight), 0), 1)
    df["forecast_daily_sales"] = np.select(
        [
            df["period_daily_sales"].gt(0) & df["history_daily_sales"].gt(0),
            df["period_daily_sales"].gt(0),
            df["history_daily_sales"].gt(0),
        ],
        [
            df["period_daily_sales"] * (1 - weight) + df["history_daily_sales"] * weight,
            df["period_daily_sales"],
            df["history_daily_sales"],
        ],
        default=0,
    )
    df["forecast_period_qty"] = df["forecast_daily_sales"] * config.period_days
    df["purchase_horizon_days"] = float(config.purchase_horizon_days)
    df["horizon_demand_qty"] = df["forecast_daily_sales"] * df["purchase_horizon_days"]
    df["unit_purchase_cost"] = np.select(
        [df["sold_qty"].gt(0), df["history_unit_purchase_cost"].gt(0)],
        [df["purchase_cost"] / df["sold_qty"].replace(0, np.nan), df["history_unit_purchase_cost"]],
        default=0,
    )
    df["unit_purchase_cost"] = pd.to_numeric(df["unit_purchase_cost"], errors="coerce").fillna(0)

    df["category_metric"] = df[metric_column].fillna(0)
    df = _assign_abcd(df, config)

    df["target_cover_days"] = float(config.purchase_horizon_days)
    df["lead_time_days"] = float(config.lead_time_days)
    df["safety_stock_days"] = float(config.safety_stock_days)
    df["stock_cover_days"] = np.where(
        df["forecast_daily_sales"].gt(0),
        df["current_stock"] / df["forecast_daily_sales"],
        np.nan,
    )
    df["safety_stock_qty"] = df["forecast_daily_sales"] * df["safety_stock_days"]
    df["target_stock"] = df["horizon_demand_qty"] + df["safety_stock_qty"]
    df["reorder_point"] = df["forecast_daily_sales"] * (
        df["lead_time_days"] + df["safety_stock_days"]
    )
    df["raw_order_qty"] = (df["target_stock"] - df["current_stock"]).clip(lower=0)
    if config.only_below_reorder_point:
        df.loc[df["current_stock"].gt(df["reorder_point"]), "raw_order_qty"] = 0

    multiple = max(int(config.order_multiple), 1)
    min_order_qty = max(int(config.min_order_qty), 0)
    rounded = np.ceil(df["raw_order_qty"] / multiple) * multiple
    rounded = np.where((rounded > 0) & (rounded < min_order_qty), min_order_qty, rounded)
    df["order_qty"] = pd.Series(rounded).fillna(0).astype(int)
    df["order_cost_estimate"] = df["order_qty"] * df["unit_purchase_cost"]
    df["stock_status"] = df.apply(_stock_status, axis=1)
    df["data_source"] = df.apply(_data_source, axis=1)

    output_columns = [
        "market",
        "seedbank",
        "product",
        "pack",
        "sku",
        "abc_category",
        "stock_status",
        "data_source",
        "sold_qty",
        "revenue",
        "margin",
        "purchase_cost",
        "current_stock",
        "period_daily_sales",
        "history_daily_sales",
        "forecast_daily_sales",
        "forecast_period_qty",
        "purchase_horizon_days",
        "horizon_demand_qty",
        "stock_cover_days",
        "target_cover_days",
        "safety_stock_qty",
        "target_stock",
        "reorder_point",
        "raw_order_qty",
        "order_qty",
        "unit_purchase_cost",
        "order_cost_estimate",
        "category_metric",
        "category_share",
        "category_cum_share",
        "lead_time_days",
        "safety_stock_days",
        "history_runs",
        "history_last_category",
        "history_last_seen_at",
    ]
    for column in output_columns:
        if column not in df.columns:
            df[column] = 0 if column in NUMERIC_COLUMNS else ""

    return (
        df[output_columns]
        .sort_values(
            by=["order_qty", "order_cost_estimate", "category_metric"],
            ascending=[False, False, False],
        )
        .reset_index(drop=True)
    )


def _assign_abcd(df: pd.DataFrame, config: AnalysisConfig) -> pd.DataFrame:
    groups = []
    for _, group in df.groupby("market", sort=False):
        group = group.sort_values("category_metric", ascending=False).copy()
        positive = group["category_metric"].gt(0)
        total = group.loc[positive, "category_metric"].sum()
        if total <= 0:
            group["category_share"] = 0.0
            group["category_cum_share"] = 0.0
            group["abc_category"] = "D"
        else:
            group["category_share"] = np.where(positive, group["category_metric"] / total, 0)
            group["category_cum_share"] = group["category_share"].cumsum()
            previous_cum_share = group["category_cum_share"] - group["category_share"]
            group["abc_category"] = np.select(
                [
                    ~positive,
                    previous_cum_share < config.threshold_a,
                    previous_cum_share < config.threshold_b,
                    previous_cum_share < config.threshold_c,
                ],
                ["D", "A", "B", "C"],
                default="D",
            )
        groups.append(group)
    return pd.concat(groups, ignore_index=True)


def _stock_status(row: pd.Series) -> str:
    if row["forecast_daily_sales"] <= 0:
        return "no_demand"
    if row["current_stock"] <= 0:
        return "out_of_stock"
    if pd.notna(row["stock_cover_days"]) and row["stock_cover_days"] <= (
        row["lead_time_days"] + row["safety_stock_days"]
    ):
        return "urgent"
    if row["order_qty"] > 0:
        return "planned"
    return "covered"


def _data_source(row: pd.Series) -> str:
    sources = []
    if bool(row.get("has_sales_row", False)):
        sources.append("current_sales")
    if bool(row.get("has_stock_row", False)):
        sources.append("current_stock")
    if bool(row.get("has_history_row", False)) or row.get("history_runs", 0) > 0:
        sources.append("history")
    return "+".join(sources) if sources else "unknown"


def summarize_result(result: pd.DataFrame) -> dict[str, float]:
    recommended = result[result["order_qty"].gt(0)]
    return {
        "items_count": float(len(result)),
        "recommended_items_count": float(len(recommended)),
        "total_sales_qty": float(result["sold_qty"].sum()),
        "total_revenue": float(result["revenue"].sum()),
        "total_margin": float(result["margin"].sum()),
        "total_order_qty": float(recommended["order_qty"].sum()),
        "total_order_cost_estimate": float(recommended["order_cost_estimate"].sum()),
        "urgent_items_count": float(result["stock_status"].isin(["out_of_stock", "urgent"]).sum()),
    }


def save_analysis(
    conn: sqlite3.Connection,
    name: str,
    config: AnalysisConfig,
    result: pd.DataFrame,
    notes: str = "",
) -> str:
    init_db(conn)
    run_uuid = uuid.uuid4().hex
    created_at = datetime.now(timezone.utc).isoformat(timespec="seconds")
    summary = summarize_result(result)
    with conn:
        conn.execute(
            """
            INSERT INTO analysis_runs (
                run_uuid,
                created_at,
                name,
                period_start,
                period_end,
                period_days,
                config_json,
                notes,
                items_count,
                recommended_items_count,
                total_sales_qty,
                total_revenue,
                total_margin,
                total_order_qty,
                total_order_cost_estimate
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_uuid,
                created_at,
                name.strip() or f"Analysis {created_at}",
                config.period_start.isoformat(),
                config.period_end.isoformat(),
                config.period_days,
                config.to_json(),
                notes,
                int(summary["items_count"]),
                int(summary["recommended_items_count"]),
                summary["total_sales_qty"],
                summary["total_revenue"],
                summary["total_margin"],
                summary["total_order_qty"],
                summary["total_order_cost_estimate"],
            ),
        )
        items = result.copy()
        items.insert(0, "created_at", created_at)
        items.insert(0, "run_uuid", run_uuid)
        item_columns = [
            "run_uuid",
            "created_at",
            "market",
            "seedbank",
            "product",
            "pack",
            "sku",
            "sold_qty",
            "purchase_cost",
            "revenue",
            "margin",
            "current_stock",
            "period_daily_sales",
            "history_daily_sales",
            "forecast_daily_sales",
            "forecast_period_qty",
            "purchase_horizon_days",
            "horizon_demand_qty",
            "unit_purchase_cost",
            "abc_category",
            "category_metric",
            "category_share",
            "category_cum_share",
            "target_cover_days",
            "lead_time_days",
            "safety_stock_days",
            "safety_stock_qty",
            "stock_cover_days",
            "target_stock",
            "reorder_point",
            "raw_order_qty",
            "order_qty",
            "order_cost_estimate",
            "stock_status",
            "data_source",
            "history_runs",
            "history_last_category",
            "history_last_seen_at",
        ]
        items[item_columns].to_sql("analysis_items", conn, if_exists="append", index=False)
    return run_uuid


def list_runs(conn: sqlite3.Connection, limit: int = 50) -> pd.DataFrame:
    init_db(conn)
    return pd.read_sql_query(
        """
        SELECT
            run_uuid,
            created_at,
            name,
            period_start,
            period_end,
            period_days,
            items_count,
            recommended_items_count,
            total_sales_qty,
            total_revenue,
            total_margin,
            total_order_qty,
            total_order_cost_estimate
        FROM analysis_runs
        ORDER BY created_at DESC
        LIMIT ?
        """,
        conn,
        params=(limit,),
    )


def get_run_items(conn: sqlite3.Connection, run_uuid: str) -> pd.DataFrame:
    init_db(conn)
    return pd.read_sql_query(
        "SELECT * FROM analysis_items WHERE run_uuid = ? ORDER BY order_qty DESC, order_cost_estimate DESC",
        conn,
        params=(run_uuid,),
    )


def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def dataframe_to_xlsx_bytes(df: pd.DataFrame, summary: dict[str, Any] | None = None) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if summary:
            summary_df = pd.DataFrame(
                [{"metric": key, "value": value} for key, value in summary.items()]
            )
            summary_df.to_excel(writer, sheet_name="summary", index=False)
        df.to_excel(writer, sheet_name="purchase_recommendation", index=False)
    return output.getvalue()


def load_local_sample_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    missing = [str(path) for path in LOCAL_SAMPLE_FILES.values() if not path.exists()]
    if missing:
        raise FileNotFoundError("Missing local sample files: " + "; ".join(missing))
    return load_current_data(
        LOCAL_SAMPLE_FILES["usa_sales"],
        LOCAL_SAMPLE_FILES["ww_sales"],
        LOCAL_SAMPLE_FILES["usa_stock"],
        LOCAL_SAMPLE_FILES["ww_stock"],
    )
