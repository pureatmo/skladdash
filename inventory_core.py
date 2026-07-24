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
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

import numpy as np
import pandas as pd


KEY_COLUMNS = ["market", "seedbank", "product", "pack"]

MONTH_NAMES_RU = {
    1: "Январь",
    2: "Февраль",
    3: "Март",
    4: "Апрель",
    5: "Май",
    6: "Июнь",
    7: "Июль",
    8: "Август",
    9: "Сентябрь",
    10: "Октябрь",
    11: "Ноябрь",
    12: "Декабрь",
}

MONTH_ALIASES = {
    1: ["Январь", "Янв", "January", "Jan"],
    2: ["Февраль", "Фев", "February", "Feb"],
    3: ["Март", "Мар", "March", "Mar"],
    4: ["Апрель", "Апр", "April", "Apr"],
    5: ["Май", "May"],
    6: ["Июнь", "Июн", "June", "Jun"],
    7: ["Июль", "Июл", "July", "Jul"],
    8: ["Август", "Авг", "August", "Aug"],
    9: ["Сентябрь", "Сен", "September", "Sep"],
    10: ["Октябрь", "Окт", "October", "Oct"],
    11: ["Ноябрь", "Ноя", "November", "Nov"],
    12: ["Декабрь", "Дек", "December", "Dec"],
}

SEASONALITY_MARKETS = ["USA", "WW", "USA+WW"]

DEFAULT_SEASONALITY_ROWS = [
    {"month": 1, "USA": 0.893654, "WW": 0.806470, "USA+WW": 0.854106},
    {"month": 2, "USA": 1.118693, "WW": 0.920489, "USA+WW": 1.028495},
    {"month": 3, "USA": 1.644017, "WW": 1.220551, "USA+WW": 1.451309},
    {"month": 4, "USA": 1.593157, "WW": 1.684132, "USA+WW": 1.634557},
    {"month": 5, "USA": 1.441737, "WW": 1.124331, "USA+WW": 1.297294},
    {"month": 6, "USA": 0.935457, "WW": 0.869876, "USA+WW": 0.905613},
    {"month": 7, "USA": 0.650268, "WW": 0.821209, "USA+WW": 0.728059},
    {"month": 8, "USA": 0.692535, "WW": 0.916595, "USA+WW": 0.794499},
    {"month": 9, "USA": 0.662112, "WW": 0.972770, "USA+WW": 0.803484},
    {"month": 10, "USA": 0.758956, "WW": 1.060091, "USA+WW": 0.895868},
    {"month": 11, "USA": 0.822589, "WW": 0.888230, "USA+WW": 0.852460},
    {"month": 12, "USA": 0.786824, "WW": 0.715256, "USA+WW": 0.754255},
]

LOCAL_SAMPLE_FILES = {
    "usa_sales": Path(r"C:\Users\Admin\Downloads\продажи USA с 01.01.2026 по 14.07.2026.xlsx"),
    "ww_sales": Path(r"C:\Users\Admin\Downloads\продажи WW с 01.01.2026 по 14.07.2026.xlsx"),
    "usa_stock": Path(r"C:\Users\Admin\Downloads\Остатки us.xlsx"),
    "ww_stock": Path(r"C:\Users\Admin\Downloads\Остатки ww.xlsx"),
}

LOCAL_SEASONALITY_FILE = Path(r"C:\Users\Admin\Downloads\Сезонноть.xlsx")

LOCAL_SAMPLE_PATTERNS = {
    "usa_sales": ["продажи USA*.xlsx", "продажи usa*.xlsx", "продажи сша*.xlsx"],
    "ww_sales": ["продажи WW*.xlsx", "продажи ww*.xlsx"],
    "usa_stock": ["Остатки us.xlsx", "Остатки US.xlsx", "US остатки*.xlsx"],
    "ww_stock": ["Остатки ww.xlsx", "Остатки WW.xlsx", "WW остатки*.xlsx"],
}

SALES_COLUMN_SYNONYMS = {
    "seedbank": ["Сидбанк", "Seedbank", "Поставщик", "Бренд", "Brand"],
    "product": ["Продукт", "Product", "Товар", "Название товара", "Name"],
    "pack": ["Упаковка", "Pack", "Фасовка", "Размер упаковки"],
    "sold_qty": ["Количество", "Кол-во", "Колво", "Продажи", "Продано", "Quantity", "Qty", "Sold Qty"],
    "purchase_cost": [
        "Закупочная стоимость",
        "Закупочная стоимость товаров",
        "Себестоимость",
        "Cost",
        "Purchase Cost",
        "Cost of Goods",
    ],
    "unit_purchase_cost": [
        "Средняя закупочная цена",
        "Средняя закупочная цена за ед",
        "Средняя закупочная цена за единицу",
        "Закупочная цена",
        "Цена закупки",
        "Purchase Price",
        "Unit Purchase Cost",
        "Cost Price",
    ],
    "revenue": ["Стоимость товаров", "Выручка", "Сумма продаж", "Revenue", "Sales Amount", "Total"],
    "unit_retail_price": ["Розничная цена", "Цена продажи", "Цена", "Retail Price", "Unit Price", "Price"],
    "margin": ["Маржа", "Прибыль", "Profit", "Gross Profit"],
    "margin_pct_source": [
        "Маржинальность",
        "Маржинальность %",
        "Profitability",
        "Margin %",
        "Margin Percent",
    ],
    "income_share_source": [
        "Доля дохода",
        "Доля от стоимости товаров",
        "Доля выручки",
        "Income Share",
        "Revenue Share",
    ],
}

REQUIRED_SALES_FIELDS = ["seedbank", "product", "pack", "sold_qty"]

SALES_DATE_COLUMNS = [
    "Дата",
    "Дата продажи",
    "Дата заказа",
    "Created at",
    "Created At",
    "Order date",
    "Order Date",
    "Date",
]

STOCK_COLUMNS = {
    "seedbank": ["Seedbank", "Сидбанк", "Поставщик", "Бренд", "Brand"],
    "product": ["Product", "Продукт", "Товар", "Название товара", "Name"],
    "pack": ["Pack", "Упаковка", "Фасовка", "Размер упаковки"],
    "current_stock": [
        "Available Qnt.",
        "Available Qty",
        "Available Quantity",
        "Свободный остаток",
        "Остаток",
        "Остатки",
        "Доступно",
        "В наличии",
        "Stock",
        "Qty",
    ],
}

REQUIRED_STOCK_FIELDS = ["seedbank", "product", "pack", "current_stock"]

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
    "base_forecast_daily_sales",
    "seasonality_factor",
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
    history_weight: float = 0.0
    history_stat: str = "p75"
    use_history_for_zero_sales: bool = False
    include_history_only: bool = False
    only_below_reorder_point: bool = False
    seasonality_months: list[int] = field(default_factory=list)
    seasonality_factors: dict[str, float] = field(default_factory=dict)

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
    _ensure_column(conn, "analysis_items", "base_forecast_daily_sales", "REAL NOT NULL DEFAULT 0")
    _ensure_column(conn, "analysis_items", "seasonality_factor", "REAL NOT NULL DEFAULT 1")
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


def _normalize_column_name(value: object) -> str:
    return re.sub(r"[^0-9a-zа-я]+", "", str(value).casefold().replace("ё", "е"))


def _find_column(columns: pd.Index, aliases: list[str]) -> str | None:
    normalized_columns = {_normalize_column_name(column): str(column) for column in columns}
    for alias in aliases:
        column = normalized_columns.get(_normalize_column_name(alias))
        if column is not None:
            return column
    return None


def _resolve_column_map(
    df: pd.DataFrame,
    synonyms: dict[str, list[str]],
    required_fields: list[str],
    source_label: str,
) -> dict[str, str]:
    column_map = {
        field: column
        for field, aliases in synonyms.items()
        if (column := _find_column(df.columns, aliases)) is not None
    }
    missing = [field for field in required_fields if field not in column_map]
    if missing:
        available = ", ".join(str(column) for column in df.columns)
        raise ValueError(
            f"{source_label}: не удалось определить обязательные поля: {', '.join(missing)}. "
            f"Доступные столбцы: {available}"
        )
    return column_map


def _find_sales_date_column(columns: pd.Index) -> str | None:
    return _find_column(columns, SALES_DATE_COLUMNS)


def _extract_sales_dates(raw: pd.DataFrame) -> pd.Series | None:
    date_column = _find_sales_date_column(raw.columns)
    if date_column is not None:
        return _parse_sales_dates(raw[date_column])

    year_column = _find_column(raw.columns, ["created_at Year", "created_at — Year", "Year", "Год"])
    month_column = _find_column(raw.columns, ["created_at Month", "created_at — Month", "Month", "Месяц"])
    day_column = _find_column(raw.columns, ["created_at Day", "created_at — Day", "Day", "День"])
    if not (year_column and month_column and day_column):
        return None

    year_text = pd.to_numeric(raw[year_column], errors="coerce").astype("Int64").astype(str).replace("<NA>", "")
    day_text = pd.to_numeric(raw[day_column], errors="coerce").astype("Int64").astype(str).replace("<NA>", "")
    date_text = year_text + " " + raw[month_column].astype(str).str.strip() + " " + day_text
    parsed = pd.to_datetime(date_text, errors="coerce", format="%Y %B %d")
    if parsed.notna().any():
        return parsed
    return pd.to_datetime(date_text, errors="coerce")


def _parse_sales_dates(series: pd.Series) -> pd.Series:
    sample = series.dropna().astype(str).str.strip().head(25)
    looks_day_first = sample.str.contains(r"^\d{1,2}[./]\d{1,2}[./]\d{2,4}", regex=True).any()
    if looks_day_first:
        return pd.to_datetime(series, errors="coerce", dayfirst=True)

    parsed = pd.to_datetime(series, errors="coerce")
    day_first = pd.to_datetime(series, errors="coerce", dayfirst=True)
    if day_first.notna().sum() > parsed.notna().sum():
        return day_first
    return parsed


def _to_number(series: pd.Series | int | float, index: pd.Index) -> pd.Series:
    if isinstance(series, pd.Series):
        source = series
    else:
        source = pd.Series(series, index=index)
    if pd.api.types.is_numeric_dtype(source):
        return pd.to_numeric(source, errors="coerce").fillna(0)
    cleaned = (
        source.fillna("")
        .astype(str)
        .str.replace("\u00a0", "", regex=False)
        .str.replace(r"\s+", "", regex=True)
        .str.replace("%", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0)


def _parse_pack_size(pack: str) -> int | None:
    numbers = [int(value) for value in re.findall(r"\d+", str(pack))]
    if not numbers:
        return None
    if "+" in str(pack):
        return sum(numbers)
    return numbers[0]


def _month_number(value: object) -> int | None:
    if pd.isna(value):
        return None
    if isinstance(value, (int, np.integer)) and 1 <= int(value) <= 12:
        return int(value)
    if isinstance(value, float) and value.is_integer() and 1 <= int(value) <= 12:
        return int(value)

    text = str(value).strip()
    if text.isdigit() and 1 <= int(text) <= 12:
        return int(text)

    normalized = _normalize_column_name(text)
    for month, aliases in MONTH_ALIASES.items():
        if normalized in {_normalize_column_name(alias) for alias in aliases}:
            return month
    return None


def _select_seasonality_sheet(sheet_names: list[str]) -> str:
    normalized = {_normalize_column_name(name): name for name in sheet_names}
    for preferred in ["Итог сезонность", "Сезонность", "Seasonality"]:
        preferred_norm = _normalize_column_name(preferred)
        if preferred_norm in normalized:
            return normalized[preferred_norm]
    for name in sheet_names:
        name_norm = _normalize_column_name(name)
        if "сезон" in name_norm or "season" in name_norm:
            return name
    return sheet_names[0]


def _coefficient_value(value: object) -> float | None:
    parsed = _to_number(pd.Series([value]), pd.Index([0])).iloc[0]
    if pd.isna(parsed) or parsed <= 0:
        return None
    if parsed > 10:
        parsed = parsed / 100
    return float(parsed)


def read_seasonality_coefficients(source: str | Path | BinaryIO) -> pd.DataFrame:
    if hasattr(source, "seek"):
        source.seek(0)
    workbook = pd.ExcelFile(source)
    sheet_name = _select_seasonality_sheet(workbook.sheet_names)
    raw = pd.read_excel(workbook, sheet_name=sheet_name, header=None)

    metric_aliases = {_normalize_column_name(alias) for alias in ["Продажи товаров", "Sold units", "Sales units"]}
    section_row = section_col = None
    for row_idx in range(raw.shape[0]):
        for col_idx in range(raw.shape[1]):
            if _normalize_column_name(raw.iat[row_idx, col_idx]) in metric_aliases:
                section_row, section_col = row_idx, col_idx
                break
        if section_row is not None:
            break

    if section_row is None or section_col is None:
        raise ValueError("В файле сезонности не найден блок `Продажи товаров`.")

    section_end_col = raw.shape[1]
    for col_idx in range(section_col + 1, raw.shape[1]):
        value = raw.iat[section_row, col_idx]
        if pd.notna(value) and _normalize_column_name(value):
            section_end_col = col_idx
            break

    market_columns: dict[str, int] = {}
    for header_row in range(section_row + 1, min(section_row + 4, raw.shape[0])):
        for col_idx in range(section_col + 1, section_end_col):
            header = _normalize_column_name(raw.iat[header_row, col_idx])
            if header in {"usa", "us"}:
                market_columns["USA"] = col_idx
            elif header == "ww":
                market_columns["WW"] = col_idx
            elif header in {"usaww", "usaдww", "usaandww"}:
                market_columns["USA+WW"] = col_idx
        if "USA" in market_columns and "WW" in market_columns:
            break

    if not market_columns:
        raise ValueError("В блоке `Продажи товаров` не найдены колонки USA/WW.")

    records: list[dict[str, object]] = []
    for row_idx in range(section_row + 1, raw.shape[0]):
        month = _month_number(raw.iat[row_idx, section_col])
        if month is None:
            continue
        record: dict[str, object] = {"month": month, "month_name": MONTH_NAMES_RU[month]}
        for market, col_idx in market_columns.items():
            record[market] = _coefficient_value(raw.iat[row_idx, col_idx])
        records.append(record)
        if len({record["month"] for record in records}) >= 12:
            break

    if not records:
        raise ValueError("В блоке `Продажи товаров` не найдены строки месяцев.")

    result = pd.DataFrame(records).drop_duplicates("month", keep="first").sort_values("month")
    for market in SEASONALITY_MARKETS:
        if market not in result.columns:
            result[market] = np.nan
        result[market] = pd.to_numeric(result[market], errors="coerce")

    result["USA+WW"] = result["USA+WW"].fillna(result[["USA", "WW"]].mean(axis=1))
    return result[["month", "month_name", *SEASONALITY_MARKETS]].reset_index(drop=True)


def average_seasonality_factors(coefficients: pd.DataFrame, months: list[int]) -> dict[str, float]:
    selected_months = [int(month) for month in months if 1 <= int(month) <= 12]
    if coefficients.empty or not selected_months:
        return {}

    selected = coefficients[coefficients["month"].isin(selected_months)].copy()
    if selected.empty:
        return {}

    factors: dict[str, float] = {}
    fallback = float(pd.to_numeric(selected.get("USA+WW"), errors="coerce").mean())
    for market in ["USA", "WW"]:
        values = pd.to_numeric(selected.get(market), errors="coerce")
        factor = float(values.mean()) if values.notna().any() else fallback
        if pd.notna(factor) and factor > 0:
            factors[market] = factor
    if pd.notna(fallback) and fallback > 0:
        factors["USA+WW"] = fallback
    return factors


def default_seasonality_coefficients() -> pd.DataFrame:
    coefficients = pd.DataFrame(DEFAULT_SEASONALITY_ROWS).copy()
    coefficients["month_name"] = coefficients["month"].map(MONTH_NAMES_RU)
    return coefficients[["month", "month_name", *SEASONALITY_MARKETS]]


def read_sales(
    source: str | Path | BinaryIO,
    market: str,
    period_start: date | None = None,
    period_end: date | None = None,
) -> pd.DataFrame:
    raw = _read_excel(source, preferred_sheet="Export")
    return _prepare_sales(raw, market, period_start, period_end)


def _prepare_sales(
    raw: pd.DataFrame,
    market: str,
    period_start: date | None = None,
    period_end: date | None = None,
) -> pd.DataFrame:
    parsed_dates = _extract_sales_dates(raw)
    if parsed_dates is not None and parsed_dates.notna().any() and period_start and period_end:
        period_start_ts = pd.Timestamp(period_start)
        period_end_ts = pd.Timestamp(period_end)
        raw = raw.loc[parsed_dates.ge(period_start_ts) & parsed_dates.le(period_end_ts)].copy()

    column_map = _resolve_column_map(raw, SALES_COLUMN_SYNONYMS, REQUIRED_SALES_FIELDS, f"{market} sales")
    df = pd.DataFrame(index=raw.index)
    for field in ["seedbank", "product", "pack"]:
        df[field] = raw[column_map[field]]

    numeric_fields = [
        "sold_qty",
        "purchase_cost",
        "unit_purchase_cost",
        "revenue",
        "unit_retail_price",
        "margin",
        "margin_pct_source",
        "income_share_source",
    ]
    for field in numeric_fields:
        source_column = column_map.get(field)
        df[field] = _to_number(raw[source_column], raw.index) if source_column else 0

    for column in ["seedbank", "product", "pack"]:
        df[column] = _clean_text(df[column])

    df["purchase_cost"] = np.where(
        df["purchase_cost"].eq(0) & df["unit_purchase_cost"].gt(0),
        df["unit_purchase_cost"] * df["sold_qty"],
        df["purchase_cost"],
    )
    df["revenue"] = np.where(
        df["revenue"].eq(0) & df["unit_retail_price"].gt(0),
        df["unit_retail_price"] * df["sold_qty"],
        df["revenue"],
    )
    calculated_margin = df["revenue"] - df["purchase_cost"]
    df["margin"] = np.where(
        df["margin"].eq(0) & calculated_margin.ne(0),
        calculated_margin,
        df["margin"],
    )
    df["margin_pct_source"] = np.where(
        df["margin_pct_source"].gt(1.5),
        df["margin_pct_source"] / 100,
        df["margin_pct_source"],
    )
    df["margin_pct_source"] = np.where(
        df["margin_pct_source"].eq(0) & df["revenue"].gt(0),
        df["margin"] / df["revenue"],
        df["margin_pct_source"],
    )
    total_revenue = float(df["revenue"].sum())
    if total_revenue > 0:
        df["income_share_source"] = np.where(
            df["income_share_source"].eq(0),
            df["revenue"] / total_revenue,
            df["income_share_source"],
        )

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


def _webhook_dataframe_from_response(data: bytes, content_type: str) -> pd.DataFrame:
    if not data:
        raise ValueError("webhook вернул пустой ответ")

    content_type = content_type.lower()
    if data.startswith(b"PK") or "spreadsheet" in content_type or "excel" in content_type:
        return _read_excel(BytesIO(data), preferred_sheet="Export")
    if "csv" in content_type:
        return pd.read_csv(BytesIO(data))
    if "json" in content_type:
        payload = json.loads(data.decode("utf-8-sig"))
        if isinstance(payload, dict):
            for key in ["data", "items", "rows", "records", "result"]:
                if isinstance(payload.get(key), list):
                    payload = payload[key]
                    break
        if isinstance(payload, list):
            return pd.json_normalize(payload)
        if isinstance(payload, dict):
            return pd.json_normalize([payload])
        raise ValueError("webhook вернул JSON в неподдерживаемом формате")

    raise ValueError(f"webhook вернул неподдерживаемый Content-Type: {content_type or 'unknown'}")


def _read_webhook_dataframe(
    webhook_url: str,
    market: str,
    period_start: date | None = None,
    period_end: date | None = None,
    timeout_seconds: int = 120,
    source_label: str = "webhook",
) -> pd.DataFrame:
    url = webhook_url.strip()
    if not url:
        raise ValueError(f"не указан URL для {source_label}")

    payload = {
        "market": market.upper(),
        "period_start": period_start.isoformat() if period_start else None,
        "period_end": period_end.isoformat() if period_end else None,
    }
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = Request(
        url,
        data=body,
        headers={
            "Content-Type": "application/json",
            "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/json,text/csv,*/*",
        },
        method="POST",
    )
    try:
        with urlopen(request, timeout=timeout_seconds) as response:
            data = response.read()
            content_type = response.headers.get("Content-Type", "")
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise ValueError(f"{source_label} вернул HTTP {exc.code}: {detail}") from exc
    except URLError as exc:
        raise ValueError(f"не удалось подключиться к {source_label}: {exc.reason}") from exc

    return _webhook_dataframe_from_response(data, content_type)


def read_sales_webhook(
    webhook_url: str,
    market: str,
    period_start: date | None = None,
    period_end: date | None = None,
    timeout_seconds: int = 120,
) -> pd.DataFrame:
    raw = _read_webhook_dataframe(
        webhook_url,
        market,
        period_start,
        period_end,
        timeout_seconds,
        source_label=f"{market.upper()} sales webhook",
    )
    return _prepare_sales(raw, market, period_start, period_end)


def read_stock(source: str | Path | BinaryIO, market: str) -> pd.DataFrame:
    raw = _read_excel(source)
    return _prepare_stock(raw, market)


def _prepare_stock(raw: pd.DataFrame, market: str) -> pd.DataFrame:
    column_map = _resolve_column_map(raw, STOCK_COLUMNS, REQUIRED_STOCK_FIELDS, f"{market} stock")
    df = pd.DataFrame(index=raw.index)
    for field in ["seedbank", "product", "pack"]:
        df[field] = raw[column_map[field]]
    df["current_stock"] = _to_number(raw[column_map["current_stock"]], raw.index)

    for column in ["seedbank", "product", "pack"]:
        df[column] = _clean_text(df[column])
    df = df[df["seedbank"].ne("") & df["product"].ne("") & df["pack"].ne("")]
    df["market"] = market.upper()
    return (
        df.groupby(KEY_COLUMNS, as_index=False)
        .agg(current_stock=("current_stock", "sum"))
        .sort_values(KEY_COLUMNS)
        .reset_index(drop=True)
    )


def read_stock_webhook(
    webhook_url: str,
    market: str,
    period_start: date | None = None,
    period_end: date | None = None,
    timeout_seconds: int = 120,
) -> pd.DataFrame:
    raw = _read_webhook_dataframe(
        webhook_url,
        market,
        period_start,
        period_end,
        timeout_seconds,
        source_label=f"{market.upper()} stock webhook",
    )
    return _prepare_stock(raw, market)


def load_current_data(
    usa_sales: str | Path | BinaryIO | None,
    ww_sales: str | Path | BinaryIO | None,
    usa_stock: str | Path | BinaryIO | None,
    ww_stock: str | Path | BinaryIO | None,
    period_start: date | None = None,
    period_end: date | None = None,
    usa_sales_webhook_url: str | None = None,
    ww_sales_webhook_url: str | None = None,
    usa_stock_webhook_url: str | None = None,
    ww_stock_webhook_url: str | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    usa_sales_frame = (
        read_sales_webhook(usa_sales_webhook_url, "USA", period_start, period_end)
        if usa_sales_webhook_url
        else read_sales(usa_sales, "USA", period_start, period_end)
    )
    ww_sales_frame = (
        read_sales_webhook(ww_sales_webhook_url, "WW", period_start, period_end)
        if ww_sales_webhook_url
        else read_sales(ww_sales, "WW", period_start, period_end)
    )
    usa_stock_frame = (
        read_stock_webhook(usa_stock_webhook_url, "USA", period_start, period_end)
        if usa_stock_webhook_url
        else read_stock(usa_stock, "USA")
    )
    ww_stock_frame = (
        read_stock_webhook(ww_stock_webhook_url, "WW", period_start, period_end)
        if ww_stock_webhook_url
        else read_stock(ww_stock, "WW")
    )
    sales = pd.concat(
        [
            usa_sales_frame,
            ww_sales_frame,
        ],
        ignore_index=True,
    )
    stock = pd.concat([usa_stock_frame, ww_stock_frame], ignore_index=True)
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
    if history.empty:
        history = _empty_history()
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
    can_use_history_only = bool(config.use_history_for_zero_sales)
    df["forecast_daily_sales"] = np.select(
        [
            df["period_daily_sales"].gt(0) & df["history_daily_sales"].gt(0),
            df["period_daily_sales"].gt(0),
            can_use_history_only & df["history_daily_sales"].gt(0),
        ],
        [
            df["period_daily_sales"] * (1 - weight) + df["history_daily_sales"] * weight,
            df["period_daily_sales"],
            df["history_daily_sales"],
        ],
        default=0,
    )
    df["base_forecast_daily_sales"] = df["forecast_daily_sales"]
    seasonality_factors = {
        str(market).upper(): float(factor)
        for market, factor in (config.seasonality_factors or {}).items()
        if pd.notna(factor) and float(factor) > 0
    }
    df["seasonality_factor"] = (
        df["market"].str.upper().map(seasonality_factors).fillna(seasonality_factors.get("USA+WW", 1.0))
    )
    df["seasonality_factor"] = pd.to_numeric(df["seasonality_factor"], errors="coerce").fillna(1.0)
    df["forecast_daily_sales"] = df["base_forecast_daily_sales"] * df["seasonality_factor"]
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
        "base_forecast_daily_sales",
        "seasonality_factor",
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
        "no_demand_items_count": float(result["stock_status"].eq("no_demand").sum()),
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
            "base_forecast_daily_sales",
            "seasonality_factor",
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


def load_local_sample_data(
    period_start: date | None = None,
    period_end: date | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    files = resolve_local_sample_files()
    missing = [str(path) for path in files.values() if not path.exists()]
    if missing:
        raise FileNotFoundError("Missing local sample files: " + "; ".join(missing))
    return load_current_data(
        files["usa_sales"],
        files["ww_sales"],
        files["usa_stock"],
        files["ww_stock"],
        period_start=period_start,
        period_end=period_end,
    )


def resolve_local_sample_files(downloads_dir: str | Path = r"C:\Users\Admin\Downloads") -> dict[str, Path]:
    downloads = Path(downloads_dir)
    resolved: dict[str, Path] = {}
    for key, fallback in LOCAL_SAMPLE_FILES.items():
        candidates: list[Path] = []
        for pattern in LOCAL_SAMPLE_PATTERNS.get(key, []):
            candidates.extend(downloads.glob(pattern))
        candidates = [path for path in candidates if path.is_file()]
        resolved[key] = max(candidates, key=lambda path: path.stat().st_mtime) if candidates else fallback
    return resolved


def resolve_local_seasonality_file(downloads_dir: str | Path = r"C:\Users\Admin\Downloads") -> Path | None:
    downloads = Path(downloads_dir)
    exact = downloads / LOCAL_SEASONALITY_FILE.name
    if exact.exists():
        return exact

    candidates = [
        path
        for path in downloads.glob("*.xlsx")
        if path.is_file() and "сезон" in path.stem.casefold()
    ]
    return max(candidates, key=lambda path: path.stat().st_mtime) if candidates else None
