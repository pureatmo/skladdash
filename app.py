from __future__ import annotations

import os
from datetime import date
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from inventory_core import (
    ABC_METRICS,
    AnalysisConfig,
    LOCAL_SAMPLE_FILES,
    compute_analysis,
    dataframe_to_csv_bytes,
    dataframe_to_xlsx_bytes,
    get_history_summary,
    get_run_items,
    list_runs,
    load_current_data,
    load_local_sample_data,
    open_connection,
    save_analysis,
    summarize_result,
)


APP_DIR = Path(__file__).resolve().parent
DB_PATH = Path(os.getenv("INVENTORY_DB_PATH", APP_DIR / "data" / "inventory_history.sqlite"))

STATUS_LABELS = {
    "out_of_stock": "Нет остатка",
    "urgent": "Срочно",
    "planned": "Планово",
    "covered": "Закрыто",
    "no_demand": "Нет спроса",
}

METRIC_LABELS = {
    "revenue": "Выручка",
    "margin": "Маржа",
    "quantity": "Количество",
}

DISPLAY_COLUMNS = {
    "market": "Рынок",
    "seedbank": "Поставщик",
    "product": "Товар",
    "pack": "Упаковка",
    "abc_category": "ABCD",
    "stock_status": "Статус",
    "data_source": "Источник",
    "sold_qty": "Продажи, шт.",
    "current_stock": "Остаток, шт.",
    "period_daily_sales": "Продажи/день",
    "history_daily_sales": "История/день",
    "forecast_daily_sales": "Прогноз/день",
    "purchase_horizon_days": "Горизонт, дней",
    "horizon_demand_qty": "Спрос на горизонт, шт.",
    "stock_cover_days": "Запас, дней",
    "target_cover_days": "Закупаем вперед, дней",
    "safety_stock_qty": "Страховой запас, шт.",
    "target_stock": "Нужно иметь, шт.",
    "raw_order_qty": "Дефицит до цели, шт.",
    "order_qty": "Заказать, шт.",
    "unit_purchase_cost": "Закупка/шт.",
    "order_cost_estimate": "Сумма закупки",
    "revenue": "Выручка",
    "margin": "Маржа",
    "history_runs": "Ист. прогонов",
}


def main() -> None:
    st.set_page_config(
        page_title="ABCD закупки",
        page_icon="",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    _inject_css()

    st.title("ABCD закупки")
    _usage_guide()

    conn = open_connection(DB_PATH)
    config, run_name, notes, use_local_files, save_to_history = _sidebar()
    input_sources = _file_inputs(use_local_files)

    runs = list_runs(conn)
    if "latest_result" not in st.session_state:
        st.session_state["latest_result"] = None
        st.session_state["latest_run_uuid"] = None

    ready = all(input_sources.values())
    latest_result = None
    current_run_uuid = None
    if ready:
        with st.spinner("Расчет..."):
            sales, stock = (
                load_local_sample_data()
                if use_local_files
                else load_current_data(
                    input_sources["usa_sales"],
                    input_sources["ww_sales"],
                    input_sources["usa_stock"],
                    input_sources["ww_stock"],
                )
            )
            history = get_history_summary(conn)
            latest_result = compute_analysis(sales, stock, history, config)
            st.session_state["latest_result"] = latest_result
            st.session_state["latest_run_uuid"] = None

        st.caption(
            "Расчет обновляется автоматически, когда вы меняете период закупа вперед, "
            "страховой запас, период продаж или другие параметры."
        )
        save_clicked = st.button(
            "Сохранить текущий расчет в историю",
            type="primary",
            disabled=not save_to_history,
        )
        if save_clicked:
            current_run_uuid = save_analysis(conn, run_name, config, latest_result, notes)
            st.session_state["latest_run_uuid"] = current_run_uuid
            runs = list_runs(conn)
            st.success("Расчет сохранен в историю.")

    if latest_result is None and not runs.empty:
        latest_result = get_run_items(conn, runs.iloc[0]["run_uuid"])
        st.session_state["latest_result"] = latest_result
        st.session_state["latest_run_uuid"] = runs.iloc[0]["run_uuid"]

    if latest_result is None:
        _empty_state(runs)
        return

    _render_dashboard(latest_result, runs, current_run_uuid or st.session_state.get("latest_run_uuid"))


def _usage_guide() -> None:
    with st.expander("Как пользоваться дашбордом", expanded=True):
        st.markdown(
            """
            **Быстрый сценарий**

            1. В левой панели укажите период продаж. Он нужен, чтобы пересчитать продажи в спрос за день.
            2. Если работаете с текущими файлами из `Downloads`, оставьте включенным `Взять текущие локальные файлы`.
               Для новых файлов выключите этот переключатель и загрузите 4 Excel-файла: продажи USA, продажи WW,
               остатки USA и остатки WW.
            3. Проверьте `Закупаем вперед, дней`. По умолчанию стоит 90 дней, то есть закуп на 3 месяца вперед.
            4. Расчет обновится автоматически после изменения параметра.
            5. Если результат нужно оставить в истории, нажмите `Сохранить текущий расчет в историю`.
            6. На вкладке `Закупка` смотрите позиции, которые нужно заказать. Фильтры сверху ограничивают рынок,
               ABCD-категории, статус остатка и поиск по товару или поставщику.
            7. На вкладке `Экспорт` скачайте закупку в CSV или XLSX.

            **Как читать результат**

            - `A` - товары, которые дают основную долю выбранной метрики, по умолчанию выручки.
            - `B` и `C` - средний хвост продаж.
            - `D` - низкий вклад или нет текущего спроса.
            - `Нет остатка` и `Срочно` - позиции с самым высоким приоритетом.
            - `Планово` - позицию стоит заказать до целевого уровня запаса.
            - `Закрыто` - текущего остатка достаточно.
            - Главное поле для закупа - `Заказать, шт.`. Статус только объясняет срочность.

            **Формула заказа**

            `Заказать, шт. = прогноз продаж в день * выбранный горизонт закупа + страховой запас в штуках - текущий остаток`.
            Если включена история, прогноз смешивает текущий период и предыдущие сохраненные анализы.
            """
        )


def _sidebar() -> tuple[AnalysisConfig, str, str, bool, bool]:
    st.sidebar.header("Параметры")
    run_name = st.sidebar.text_input("Название закупа", value=f"Закуп {date.today().isoformat()}")
    period_start = st.sidebar.date_input("Период с", value=date(2026, 1, 1))
    period_end = st.sidebar.date_input("Период по", value=date(2026, 7, 14))
    if period_end < period_start:
        st.sidebar.error("Дата окончания раньше даты начала")
        period_end = period_start

    abc_metric_label = st.sidebar.selectbox(
        "ABCD по метрике",
        options=list(METRIC_LABELS.keys()),
        format_func=lambda value: METRIC_LABELS[value],
        index=0,
    )
    threshold_a = st.sidebar.slider("A до накопленной доли", 50, 95, 80, 1) / 100
    threshold_b = st.sidebar.slider("B до накопленной доли", 70, 99, 95, 1) / 100
    threshold_c = st.sidebar.slider("C до накопленной доли", 80, 100, 99, 1) / 100
    if not (threshold_a < threshold_b < threshold_c):
        st.sidebar.warning("Пороги должны расти: A < B < C")

    st.sidebar.header("Горизонт закупа")
    purchase_horizon_days = st.sidebar.number_input(
        "Закупаем вперед, дней",
        min_value=1,
        value=90,
        step=5,
        help="90 дней = закуп примерно на 3 месяца вперед. Это главный параметр количества к заказу.",
    )
    cover_days = {category: int(purchase_horizon_days) for category in ["A", "B", "C", "D"]}
    lead_time_days = st.sidebar.number_input("Срок поставки, дней", min_value=0, value=45, step=1)
    safety_stock_days = st.sidebar.number_input("Страховой запас, дней", min_value=0, value=15, step=1)
    min_order_qty = st.sidebar.number_input("Минимальный заказ, шт.", min_value=0, value=1, step=1)
    order_multiple = st.sidebar.number_input("Кратность заказа, шт.", min_value=1, value=1, step=1)
    only_below_reorder = st.sidebar.checkbox("Только ниже точки заказа", value=False)

    st.sidebar.header("История")
    history_weight = st.sidebar.slider("Вес истории в прогнозе", 0, 100, 35, 5) / 100
    history_stat = st.sidebar.selectbox(
        "Исторический спрос",
        options=["p75", "avg"],
        format_func=lambda value: "75-й перцентиль" if value == "p75" else "Среднее",
    )
    include_history_only = st.sidebar.checkbox("Добавлять SKU только из истории", value=True)
    save_to_history = st.sidebar.checkbox("Разрешить сохранение в историю", value=True)
    notes = st.sidebar.text_area("Заметки", value="", height=80)

    local_available = all(path.exists() for path in LOCAL_SAMPLE_FILES.values())
    use_local_files = st.sidebar.checkbox(
        "Взять текущие локальные файлы",
        value=local_available,
        disabled=not local_available,
    )
    st.sidebar.caption(f"DB: {DB_PATH}")

    config = AnalysisConfig(
        period_start=period_start,
        period_end=period_end,
        abc_metric=abc_metric_label,
        threshold_a=threshold_a,
        threshold_b=threshold_b,
        threshold_c=threshold_c,
        cover_days_by_category=cover_days,
        lead_time_days=int(lead_time_days),
        safety_stock_days=int(safety_stock_days),
        purchase_horizon_days=int(purchase_horizon_days),
        min_order_qty=int(min_order_qty),
        order_multiple=int(order_multiple),
        history_weight=history_weight,
        history_stat=history_stat,
        include_history_only=include_history_only,
        only_below_reorder_point=only_below_reorder,
    )
    return config, run_name, notes, use_local_files, save_to_history


def _file_inputs(use_local_files: bool) -> dict[str, object]:
    if use_local_files:
        return {key: path for key, path in LOCAL_SAMPLE_FILES.items()}

    st.subheader("Файлы")
    cols = st.columns(4)
    with cols[0]:
        usa_sales = st.file_uploader("Продажи USA", type=["xlsx"], key="usa_sales")
    with cols[1]:
        ww_sales = st.file_uploader("Продажи WW", type=["xlsx"], key="ww_sales")
    with cols[2]:
        usa_stock = st.file_uploader("Остатки USA", type=["xlsx"], key="usa_stock")
    with cols[3]:
        ww_stock = st.file_uploader("Остатки WW", type=["xlsx"], key="ww_stock")
    return {
        "usa_sales": usa_sales,
        "ww_sales": ww_sales,
        "usa_stock": usa_stock,
        "ww_stock": ww_stock,
    }


def _render_dashboard(result: pd.DataFrame, runs: pd.DataFrame, run_uuid: str | None) -> None:
    summary = summarize_result(result)
    _metrics(summary, run_uuid)

    purchase, filtered = _filters(result)
    tabs = st.tabs(["Закупка", "ABCD", "История", "Экспорт"])

    with tabs[0]:
        _purchase_tab(purchase, filtered)
    with tabs[1]:
        _abcd_tab(filtered)
    with tabs[2]:
        _history_tab(runs)
    with tabs[3]:
        _export_tab(purchase, result, summary)


def _metrics(summary: dict[str, float], run_uuid: str | None) -> None:
    cols = st.columns(6)
    cols[0].metric("SKU", f"{summary['items_count']:,.0f}")
    cols[1].metric("К заказу", f"{summary['recommended_items_count']:,.0f}")
    cols[2].metric("Штук", f"{summary['total_order_qty']:,.0f}")
    cols[3].metric("Сумма", f"{summary['total_order_cost_estimate']:,.0f}")
    cols[4].metric("Выручка", f"{summary['total_revenue']:,.0f}")
    cols[5].metric("Срочно", f"{summary['urgent_items_count']:,.0f}")
    if run_uuid:
        st.caption(f"run: {run_uuid}")


def _filters(result: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    filter_cols = st.columns([1, 1, 1, 2])
    markets = filter_cols[0].multiselect("Рынок", sorted(result["market"].unique()), default=sorted(result["market"].unique()))
    categories = filter_cols[1].multiselect(
        "ABCD", ["A", "B", "C", "D"], default=["A", "B", "C", "D"]
    )
    statuses = filter_cols[2].multiselect(
        "Статус",
        list(STATUS_LABELS.keys()),
        default=["out_of_stock", "urgent", "planned"],
        format_func=lambda value: STATUS_LABELS[value],
    )
    search = filter_cols[3].text_input("Поиск", value="")

    filtered = result[
        result["market"].isin(markets)
        & result["abc_category"].isin(categories)
        & result["stock_status"].isin(statuses)
    ].copy()
    if search.strip():
        needle = search.strip().casefold()
        haystack = (
            filtered["seedbank"].astype(str)
            + " "
            + filtered["product"].astype(str)
            + " "
            + filtered["pack"].astype(str)
        ).str.casefold()
        filtered = filtered[haystack.str.contains(needle, regex=False)]

    purchase = filtered[filtered["order_qty"].gt(0)].copy()
    return purchase, filtered


def _purchase_tab(purchase: pd.DataFrame, filtered: pd.DataFrame) -> None:
    horizon_days = 90
    if "purchase_horizon_days" in filtered.columns and not filtered.empty:
        horizon_days = int(filtered["purchase_horizon_days"].max())
    st.info(
        f"Количество `Заказать, шт.` считается на {horizon_days} дней вперед: "
        "`Прогноз/день * горизонт + страховой запас - текущий остаток`. "
        "Статус показывает срочность, но итоговое количество берется из этой формулы."
    )

    chart_cols = st.columns([1.25, 1])
    with chart_cols[0]:
        by_supplier = (
            purchase.groupby(["market", "seedbank"], as_index=False)
            .agg(order_qty=("order_qty", "sum"), order_cost_estimate=("order_cost_estimate", "sum"))
            .sort_values("order_qty", ascending=False)
            .head(30)
        )
        if not by_supplier.empty:
            fig = px.bar(
                by_supplier,
                x="order_qty",
                y="seedbank",
                color="market",
                orientation="h",
                labels={"order_qty": "Штук", "seedbank": "Поставщик", "market": "Рынок"},
                height=520,
            )
            fig.update_layout(yaxis={"categoryorder": "total ascending"}, margin=dict(l=10, r=10, t=20, b=10))
            st.plotly_chart(fig, width="stretch")
    with chart_cols[1]:
        status = (
            filtered.groupby(["stock_status"], as_index=False)
            .agg(sku=("sku", "count"), order_qty=("order_qty", "sum"))
            .sort_values("order_qty", ascending=False)
        )
        status["stock_status"] = status["stock_status"].map(STATUS_LABELS)
        if not status.empty:
            fig = px.bar(
                status,
                x="stock_status",
                y="order_qty",
                color="stock_status",
                labels={"stock_status": "Статус", "order_qty": "Штук"},
                height=520,
            )
            fig.update_layout(showlegend=False, margin=dict(l=10, r=10, t=20, b=10))
            st.plotly_chart(fig, width="stretch")

    st.dataframe(_format_table(purchase), width="stretch", height=560)


def _abcd_tab(filtered: pd.DataFrame) -> None:
    cols = st.columns([1, 1])
    with cols[0]:
        abc = (
            filtered.groupby(["market", "abc_category"], as_index=False)
            .agg(
                sku=("sku", "count"),
                sold_qty=("sold_qty", "sum"),
                revenue=("revenue", "sum"),
                order_qty=("order_qty", "sum"),
            )
            .sort_values(["market", "abc_category"])
        )
        fig = px.bar(
            abc,
            x="abc_category",
            y="revenue",
            color="market",
            barmode="group",
            labels={"abc_category": "ABCD", "revenue": "Выручка", "market": "Рынок"},
            height=430,
        )
        fig.update_layout(margin=dict(l=10, r=10, t=20, b=10))
        st.plotly_chart(fig, width="stretch")
    with cols[1]:
        fig = px.scatter(
            filtered,
            x="forecast_daily_sales",
            y="current_stock",
            color="abc_category",
            size="order_qty",
            hover_data=["market", "seedbank", "product", "pack", "stock_status"],
            labels={
                "forecast_daily_sales": "Прогноз/день",
                "current_stock": "Остаток",
                "abc_category": "ABCD",
            },
            height=430,
        )
        fig.update_layout(margin=dict(l=10, r=10, t=20, b=10))
        st.plotly_chart(fig, width="stretch")

    st.dataframe(_format_table(filtered), width="stretch", height=560)


def _history_tab(runs: pd.DataFrame) -> None:
    if runs.empty:
        st.info("История пуста")
        return
    display = runs.copy()
    display["created_at"] = pd.to_datetime(display["created_at"]).dt.strftime("%Y-%m-%d %H:%M")
    display = display.rename(
        columns={
            "created_at": "Дата",
            "name": "Название",
            "period_start": "С",
            "period_end": "По",
            "items_count": "SKU",
            "recommended_items_count": "К заказу",
            "total_sales_qty": "Продажи",
            "total_revenue": "Выручка",
            "total_margin": "Маржа",
            "total_order_qty": "Заказать",
            "total_order_cost_estimate": "Сумма закупки",
        }
    )
    st.dataframe(display.drop(columns=["run_uuid"]), width="stretch", height=360)

    trend = runs.sort_values("created_at").copy()
    if len(trend) > 1:
        trend["created_at"] = pd.to_datetime(trend["created_at"])
        fig = px.line(
            trend,
            x="created_at",
            y=["total_order_qty", "total_order_cost_estimate"],
            labels={"created_at": "Дата", "value": "Значение", "variable": "Метрика"},
            height=420,
        )
        fig.update_layout(margin=dict(l=10, r=10, t=20, b=10))
        st.plotly_chart(fig, width="stretch")


def _export_tab(purchase: pd.DataFrame, result: pd.DataFrame, summary: dict[str, float]) -> None:
    export_cols = st.columns(4)
    export_cols[0].download_button(
        "CSV закупка",
        dataframe_to_csv_bytes(purchase),
        file_name="purchase_recommendation.csv",
        mime="text/csv",
        disabled=purchase.empty,
    )
    export_cols[1].download_button(
        "XLSX закупка",
        dataframe_to_xlsx_bytes(purchase, summary),
        file_name="purchase_recommendation.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=purchase.empty,
    )
    export_cols[2].download_button(
        "CSV все SKU",
        dataframe_to_csv_bytes(result),
        file_name="inventory_abcd_all_sku.csv",
        mime="text/csv",
    )
    export_cols[3].download_button(
        "XLSX все SKU",
        dataframe_to_xlsx_bytes(result, summary),
        file_name="inventory_abcd_all_sku.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def _format_table(df: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "market",
        "seedbank",
        "product",
        "pack",
        "abc_category",
        "stock_status",
        "sold_qty",
        "current_stock",
        "period_daily_sales",
        "history_daily_sales",
        "forecast_daily_sales",
        "purchase_horizon_days",
        "horizon_demand_qty",
        "stock_cover_days",
        "safety_stock_qty",
        "target_stock",
        "raw_order_qty",
        "order_qty",
        "unit_purchase_cost",
        "order_cost_estimate",
        "revenue",
        "margin",
        "history_runs",
        "data_source",
    ]
    table = df[[column for column in columns if column in df.columns]].copy()
    table["stock_status"] = table["stock_status"].map(STATUS_LABELS).fillna(table["stock_status"])
    table = table.rename(columns=DISPLAY_COLUMNS)
    return table


def _empty_state(runs: pd.DataFrame) -> None:
    if runs.empty:
        st.info("Загрузите 4 файла и запустите анализ")
    else:
        st.info("Запустите новый анализ или откройте последний сохраненный прогон")


def _inject_css() -> None:
    st.markdown(
        """
        <style>
        .block-container { padding-top: 1.25rem; padding-bottom: 2rem; }
        h1 { font-size: 2rem; margin-bottom: .35rem; }
        div[data-testid="stMetric"] {
            background: #ffffff;
            border: 1px solid #e6e8eb;
            border-radius: 8px;
            padding: .75rem .85rem;
        }
        div[data-testid="stMetricValue"] { font-size: 1.45rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
