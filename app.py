from __future__ import annotations

import os
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from inventory_core import (
    ABC_METRICS,
    AnalysisConfig,
    compute_analysis,
    dataframe_to_csv_bytes,
    dataframe_to_xlsx_bytes,
    get_history_summary,
    get_run_items,
    list_runs,
    load_current_data,
    load_local_sample_data,
    open_connection,
    resolve_local_sample_files,
    save_analysis,
    summarize_result,
)


APP_DIR = Path(__file__).resolve().parent
DB_PATH = Path(os.getenv("INVENTORY_DB_PATH", APP_DIR / "data" / "inventory_history.sqlite"))
CRITICAL_COVER_DAYS = 14

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

COLUMN_TOOLTIPS = {
    "Рынок": "Склад/рынок, к которому относится строка: USA или WW.",
    "Поставщик": "Seedbank из файла продаж или остатков.",
    "Товар": "Название товара из файла продаж или остатков.",
    "Упаковка": "Конкретная упаковка товара. Каждая упаковка считается отдельной SKU-строкой.",
    "ABCD": "Категория по накопленной доле выбранной метрики внутри рынка: A - основной вклад, D - низкий вклад или нет спроса.",
    "Статус": "Срочность закупки по текущему запасу, прогнозу продаж, сроку поставки и страховому запасу.",
    "Источник": "Откуда взята строка: текущие продажи/остатки или только исторические данные.",
    "Продажи, шт.": "Количество, проданное в загруженном файле продаж за выбранный период продаж.",
    "Остаток, шт.": "Текущий остаток из загруженного файла остатков для этой упаковки.",
    "Продажи/день": "Продажи из файла, разделенные на количество дней в выбранном периоде продаж.",
    "История/день": "Дневной спрос по ранее сохраненным анализам: среднее или 75-й перцентиль, как выбрано в настройках.",
    "Прогноз/день": "Дневной спрос, который используется для закупа. По умолчанию равен продажам/день из текущего файла; при 0 текущих продаж история не подставляется без отдельного чекбокса.",
    "Горизонт, дней": "На сколько дней вперед рассчитывается закуп.",
    "Спрос на горизонт, шт.": "Прогноз/день, умноженный на выбранный горизонт закупа.",
    "Запас, дней": "На сколько дней хватит текущего остатка: остаток / прогноз/день.",
    "Закупаем вперед, дней": "Целевой горизонт закупа из боковой панели.",
    "Страховой запас, шт.": "Дополнительный запас в штуках: прогноз/день * страховой запас в днях.",
    "Нужно иметь, шт.": "Сколько нужно иметь на складе под горизонт закупа и страховой запас.",
    "Дефицит до цели, шт.": "Нехватка до целевого остатка до округления по минимуму и кратности заказа.",
    "Заказать, шт.": "Итоговая рекомендация к заказу: дефицит после округления по минимальному заказу и кратности.",
    "Закупка/шт.": "Расчетная закупочная цена за штуку из текущих продаж или истории.",
    "Сумма закупки": "Оценка суммы закупки: заказать, шт. * закупка/шт.",
    "Выручка": "Выручка из загруженного файла продаж за выбранный период.",
    "Маржа": "Маржа из загруженного файла продаж за выбранный период.",
    "Ист. прогонов": "Сколько сохраненных анализов использовалось для исторического спроса по этой упаковке.",
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
    config, run_name, notes, use_local_files, save_to_history, auto_sales_period = _sidebar()
    input_sources = _file_inputs(use_local_files)
    _calculation_context(config, auto_sales_period)

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
                load_local_sample_data(config.period_start, config.period_end)
                if use_local_files
                else load_current_data(
                    input_sources["usa_sales"],
                    input_sources["ww_sales"],
                    input_sources["usa_stock"],
                    input_sources["ww_stock"],
                    period_start=config.period_start,
                    period_end=config.period_end,
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

    if latest_result is None:
        _empty_state(runs)
        return

    _render_dashboard(latest_result, runs, current_run_uuid or st.session_state.get("latest_run_uuid"))


def _usage_guide() -> None:
    with st.expander("Как пользоваться дашбордом", expanded=True):
        st.markdown(
            """
            **Быстрый сценарий**

            1. В левой панели укажите горизонт закупа. По умолчанию период продаж автоматически берется такой же длины назад от последней даты продаж.
            2. В блоке `Источник данных` выберите `Загрузить новые файлы` и загрузите 4 Excel-файла:
               продажи USA, продажи WW, остатки USA и остатки WW. Режим `Downloads` нужен только для локальной проверки на этом компьютере.
            3. Проверьте `Последняя дата продаж`. Например, если закупаем на 90 дней вперед и последняя дата 22.07.2026, дашборд считает продажи за 24.04.2026-22.07.2026.
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
            - `Наличие A по упаковкам` считает долю упаковок/SKU группы A, где остаток больше 0. Количество штук внутри упаковки в этот процент не входит.

            **Формула заказа**

            `Заказать, шт. = прогноз продаж в день * выбранный горизонт закупа + страховой запас в штуках - текущий остаток`.
            В режиме авто-периода `прогноз продаж в день` считается по продажам за такое же количество дней назад, как горизонт закупа вперед.
            Если файл продаж агрегированный и в нем нет дат по строкам, выгружайте файл продаж сразу за этот период.
            История не создает закуп при нулевых текущих продажах, пока не включен отдельный чекбокс `Использовать историю при 0 продаж`.
            По умолчанию текущие продажи главнее истории. Историю можно включить отдельно для старых SKU или вручную поднять ее вес в прогнозе.
            """
        )


def _calculation_context(config: AnalysisConfig, auto_sales_period: bool) -> None:
    period_start = config.period_start.strftime("%d.%m.%Y")
    period_end = config.period_end.strftime("%d.%m.%Y")
    history_pct = int(round(config.history_weight * 100))
    period_rule = (
        f"Авто-период включен: при горизонте закупа {config.purchase_horizon_days} дней "
        f"продажи берутся за {config.purchase_horizon_days} дней назад от последней даты продаж."
        if auto_sales_period
        else "Авто-период выключен: период продаж задается вручную датами начала и окончания."
    )
    if history_pct:
        history_rule = (
            f"История смешивается с текущими продажами с весом {history_pct}%. "
            "Для SKU с нулевыми текущими продажами история используется только если включен чекбокс `Использовать историю при 0 продаж`."
        )
    else:
        history_rule = (
            "История не смешивается с текущими продажами. Если у SKU есть продажи в новом файле, "
            "прогноз считается по ним. При 0 продаж прогноз остается 0, если отдельно не включить `Использовать историю при 0 продаж`."
        )

    st.info(
        f"Период продаж для расчета: **{period_start} - {period_end}** "
        f"(**{config.period_days} дней**). "
        f"{period_rule} "
        f"`Продажи/день = Продажи, шт. / {config.period_days}`. "
        f"`Спрос на горизонт = Прогноз/день * {config.purchase_horizon_days} дней`. "
        f"`Заказать, шт. = max(0, Спрос на горизонт + Страховой запас - Остаток)`, "
        "затем применяется минимальный заказ и кратность. "
        "Если файл продаж агрегированный без дат по строкам, его нужно выгружать за этот же период; "
        "иначе старые продажи из большого файла нельзя автоматически исключить. "
        f"{history_rule}"
    )


def _sidebar() -> tuple[AnalysisConfig, str, str, bool, bool, bool]:
    st.sidebar.header("Параметры")
    run_name = st.sidebar.text_input(
        "Название закупа",
        value=f"Закуп {date.today().isoformat()}",
        help="Название сохраненного анализа в истории. На формулы не влияет.",
    )

    st.sidebar.header("Горизонт и период продаж")
    purchase_horizon_days = st.sidebar.number_input(
        "Закупаем вперед, дней",
        min_value=1,
        value=90,
        step=5,
        help="На сколько дней вперед нужен запас. При включенном авто-периоде продажи анализируются за столько же дней назад.",
    )
    auto_sales_period = st.sidebar.checkbox(
        "Период продаж = горизонту закупа назад",
        value=True,
        help="Если включено, дата начала продаж считается автоматически: последняя дата продаж минус выбранный горизонт закупа плюс 1 день.",
    )
    period_end = st.sidebar.date_input(
        "Последняя дата продаж",
        value=date.today(),
        help="Последний день продаж в загруженном отчете. Авто-период отсчитывается назад от этой даты.",
    )
    if auto_sales_period:
        period_start = period_end - timedelta(days=int(purchase_horizon_days) - 1)
        st.sidebar.caption(
            f"Период продаж для расчета: {period_start.strftime('%d.%m.%Y')} - "
            f"{period_end.strftime('%d.%m.%Y')} ({int(purchase_horizon_days)} дней)."
        )
    else:
        period_start = st.sidebar.date_input(
            "Период продаж с",
            value=period_end - timedelta(days=89),
            help="Дата начала периода, за который сформирован загруженный файл продаж.",
        )
    if period_end < period_start:
        st.sidebar.error("Дата окончания раньше даты начала")
        period_end = period_start

    abc_metric_label = st.sidebar.selectbox(
        "ABCD по метрике",
        options=list(METRIC_LABELS.keys()),
        format_func=lambda value: METRIC_LABELS[value],
        index=0,
        help="По какой метрике сортировать товары для ABCD: выручка, маржа или количество продаж.",
    )
    threshold_a = st.sidebar.slider(
        "A до накопленной доли",
        50,
        95,
        80,
        1,
        help="Граница группы A по накопленной доле выбранной метрики. 80 значит: товары, которые вместе дают первые 80% выручки/маржи/количества.",
    ) / 100
    threshold_b = st.sidebar.slider(
        "B до накопленной доли",
        70,
        99,
        95,
        1,
        help="Граница группы B по накопленной доле. При 95 группа B идет после A до 95% общего вклада.",
    ) / 100
    threshold_c = st.sidebar.slider(
        "C до накопленной доли",
        80,
        100,
        99,
        1,
        help="Граница группы C по накопленной доле. Все, что ниже этой границы, попадает в C; остальное в D.",
    ) / 100
    if not (threshold_a < threshold_b < threshold_c):
        st.sidebar.warning("Пороги должны расти: A < B < C")

    cover_days = {category: int(purchase_horizon_days) for category in ["A", "B", "C", "D"]}
    lead_time_days = st.sidebar.number_input(
        "Срок поставки, дней",
        min_value=0,
        value=45,
        step=1,
        help="Используется для статуса срочности и точки заказа: прогноз/день * (срок поставки + страховой запас).",
    )
    safety_stock_days = st.sidebar.number_input(
        "Страховой запас, дней",
        min_value=0,
        value=15,
        step=1,
        help="Дополнительный запас сверх горизонта закупа. В штуках считается как прогноз/день * страховой запас, дней.",
    )
    min_order_qty = st.sidebar.number_input(
        "Минимальный заказ, шт.",
        min_value=0,
        value=1,
        step=1,
        help="Если расчет дал ненулевой заказ меньше этого значения, рекомендация поднимется до минимума.",
    )
    order_multiple = st.sidebar.number_input(
        "Кратность заказа, шт.",
        min_value=1,
        value=1,
        step=1,
        help="Итоговый заказ округляется вверх до кратности. Например, кратность 5 округлит 12 до 15.",
    )
    only_below_reorder = st.sidebar.checkbox(
        "Только ниже точки заказа",
        value=False,
        help="Если включено, дашборд рекомендует закуп только когда текущий остаток ниже точки заказа: прогноз/день * (срок поставки + страховой запас).",
    )

    st.sidebar.header("История")
    history_weight = st.sidebar.slider(
        "Вес истории для SKU с текущими продажами",
        0,
        100,
        0,
        5,
        key="history_weight_current_sales",
        help="0% значит: если в новом файле есть продажи, прогноз считается только по ним. На SKU с нулевыми продажами влияет только при включенном чекбоксе ниже.",
    ) / 100
    history_stat = st.sidebar.selectbox(
        "Исторический спрос",
        options=["p75", "avg"],
        format_func=lambda value: "75-й перцентиль" if value == "p75" else "Среднее",
        help="Как считать дневной спрос из сохраненной истории: среднее мягче, 75-й перцентиль закладывает более высокий спрос.",
    )
    include_history_only = st.sidebar.checkbox(
        "Добавлять SKU только из истории",
        value=False,
        key="include_history_only_sku",
        help="Включайте только если хотите предлагать к закупке позиции, которых нет в новых продажах или остатках. Иначе старые pack-версии могут попадать в закупку.",
    )
    use_history_for_zero_sales = st.sidebar.checkbox(
        "Использовать историю при 0 продаж",
        value=False,
        key="use_history_for_zero_sales",
        help="Если выключено, SKU с продажами 0 в текущем файле получит прогноз 0 и не попадет в закуп только из-за истории. Включайте для новых/возвращаемых SKU, где текущий период продаж не показателен.",
    )
    save_to_history = st.sidebar.checkbox(
        "Разрешить сохранение в историю",
        value=True,
        help="Если выключить, кнопку сохранения текущего расчета нельзя будет нажать.",
    )
    notes = st.sidebar.text_area(
        "Заметки",
        value="",
        height=80,
        help="Комментарий к закупке или источнику данных. Сохраняется вместе с прогоном.",
    )

    st.sidebar.header("Источник данных")
    local_files = resolve_local_sample_files()
    local_available = all(path.exists() for path in local_files.values())
    source_options = ["Загрузить новые файлы"]
    if local_available:
        source_options.append("Взять локальные файлы из Downloads")
    source_choice = st.sidebar.radio(
        "Откуда брать Excel",
        source_options,
        index=0,
        key="data_source_choice",
        help="Для онлайн-версии используйте загрузку новых файлов. Локальные файлы работают только на этом компьютере.",
    )
    use_local_files = source_choice == "Взять локальные файлы из Downloads"
    if use_local_files:
        st.sidebar.caption(
            "Локальные файлы: "
            + "; ".join(f"{name}: {path.name}" for name, path in local_files.items())
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
        use_history_for_zero_sales=use_history_for_zero_sales,
        include_history_only=include_history_only,
        only_below_reorder_point=only_below_reorder,
    )
    return config, run_name, notes, use_local_files, save_to_history, auto_sales_period


def _file_inputs(use_local_files: bool) -> dict[str, object]:
    if use_local_files:
        return {key: path for key, path in resolve_local_sample_files().items()}

    st.subheader("Загрузка новых файлов")
    st.caption(
        "Загрузите 4 Excel-файла. После загрузки расчет обновится автоматически. "
        "Если файл продаж без дат по строкам, выгружайте его сразу за период продаж, показанный выше."
    )
    cols = st.columns(4)
    with cols[0]:
        usa_sales = st.file_uploader(
            "Продажи USA",
            type=["xlsx"],
            key="usa_sales",
            help="Файл продаж USA за период, который указан слева в полях `Период продаж с/по`.",
        )
    with cols[1]:
        ww_sales = st.file_uploader(
            "Продажи WW",
            type=["xlsx"],
            key="ww_sales",
            help="Файл продаж WW за период, который указан слева в полях `Период продаж с/по`.",
        )
    with cols[2]:
        usa_stock = st.file_uploader(
            "Остатки USA",
            type=["xlsx"],
            key="usa_stock",
            help="Текущий остаток USA на дату формирования закупа.",
        )
    with cols[3]:
        ww_stock = st.file_uploader(
            "Остатки WW",
            type=["xlsx"],
            key="ww_stock",
            help="Текущий остаток WW на дату формирования закупа.",
        )
    return {
        "usa_sales": usa_sales,
        "ww_sales": ww_sales,
        "usa_stock": usa_stock,
        "ww_stock": ww_stock,
    }


def _render_dashboard(result: pd.DataFrame, runs: pd.DataFrame, run_uuid: str | None) -> None:
    summary = summarize_result(result)
    _metrics(summary, run_uuid)
    _a_stock_availability_panel(result)

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
    cols[0].metric(
        "SKU",
        f"{summary['items_count']:,.0f}",
        help="Количество строк `рынок + поставщик + товар + упаковка`, попавших в текущий расчет.",
    )
    cols[1].metric(
        "К заказу",
        f"{summary['recommended_items_count']:,.0f}",
        help="Сколько SKU получили рекомендацию `Заказать, шт.` больше 0.",
    )
    cols[2].metric(
        "Штук",
        f"{summary['total_order_qty']:,.0f}",
        help="Общее количество штук к заказу по всем отфильтрованным SKU.",
    )
    cols[3].metric(
        "Сумма",
        f"{summary['total_order_cost_estimate']:,.0f}",
        help="Оценка стоимости закупа: сумма `Заказать, шт. * Закупка/шт.`.",
    )
    cols[4].metric(
        "Выручка",
        f"{summary['total_revenue']:,.0f}",
        help="Выручка из загруженных файлов продаж за выбранный период продаж.",
    )
    cols[5].metric(
        "Срочно",
        f"{summary['urgent_items_count']:,.0f}",
        help="Количество SKU со статусом `Нет остатка` или `Срочно`.",
    )
    if run_uuid:
        st.caption(f"run: {run_uuid}")


def _a_stock_availability_panel(result: pd.DataFrame) -> None:
    a_items = result[result["abc_category"].eq("A")].copy()
    total_packs = len(a_items)
    available_packs = int(a_items["current_stock"].gt(0).sum()) if total_packs else 0
    availability_pct = (available_packs / total_packs * 100) if total_packs else 0

    critical = a_items[
        a_items["forecast_daily_sales"].gt(0)
        & a_items["stock_cover_days"].notna()
        & a_items["stock_cover_days"].le(CRITICAL_COVER_DAYS)
    ].copy()

    with st.container(border=True):
        top_cols = st.columns([1, 1, 1])
        top_cols[0].metric(
            "Наличие A по упаковкам",
            f"{availability_pct:.1f}%",
            help="Доля упаковок/SKU группы A, где текущий остаток больше 0. Считается по строкам, не по количеству штук.",
        )
        top_cols[1].metric(
            "A-упаковок в наличии",
            f"{available_packs:,}".replace(",", " "),
            help="Количество SKU группы A с остатком больше 0.",
        )
        top_cols[2].metric(
            "A-упаковок всего",
            f"{total_packs:,}".replace(",", " "),
            help="Общее количество SKU группы A в текущем расчете.",
        )
        st.caption(
            "Процент считается по упаковкам/SKU, а не по количеству штук на складе: "
            "каждая строка `рынок + поставщик + товар + упаковка` считается как 1 упаковка."
        )

        if critical.empty:
            st.success(f"Критически низких A-упаковок с запасом {CRITICAL_COVER_DAYS} дней или меньше нет.")
            return

        st.warning(
            f"Есть {len(critical)} A-упаковок, где запас на складе {CRITICAL_COVER_DAYS} дней или меньше. "
            "Их нужно смотреть отдельно, даже если общий процент наличия выглядит нормальным."
        )
        critical_preview = (
            critical.sort_values(["stock_cover_days", "current_stock"], ascending=[True, True])
            .head(12)[
                [
                    "market",
                    "seedbank",
                    "product",
                    "pack",
                    "current_stock",
                    "forecast_daily_sales",
                    "stock_cover_days",
                    "order_qty",
                ]
            ]
            .rename(
                columns={
                    "market": "Рынок",
                    "seedbank": "Поставщик",
                    "product": "Товар",
                    "pack": "Упаковка",
                    "current_stock": "Остаток, шт.",
                    "forecast_daily_sales": "Прогноз/день",
                    "stock_cover_days": "Запас, дней",
                    "order_qty": "Заказать, шт.",
                }
            )
        )
        st.dataframe(
            critical_preview,
            width="stretch",
            height=260,
            column_config=_table_column_config(),
        )


def _filters(result: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    filter_cols = st.columns([1, 1, 1, 2])
    markets = filter_cols[0].multiselect(
        "Рынок",
        sorted(result["market"].unique()),
        default=sorted(result["market"].unique()),
        help="Ограничивает расчетную таблицу по складу/рынку.",
    )
    categories = filter_cols[1].multiselect(
        "ABCD",
        ["A", "B", "C", "D"],
        default=["A", "B", "C", "D"],
        help="Оставляет только выбранные ABCD-группы.",
    )
    statuses = filter_cols[2].multiselect(
        "Статус",
        list(STATUS_LABELS.keys()),
        default=["out_of_stock", "urgent", "planned"],
        format_func=lambda value: STATUS_LABELS[value],
        help="Фильтр по срочности закупки. На сам расчет количества не влияет.",
    )
    search = filter_cols[3].text_input(
        "Поиск",
        value="",
        help="Поиск по поставщику, названию товара и упаковке.",
    )

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

    st.dataframe(
        _format_table(purchase),
        width="stretch",
        height=560,
        column_config=_table_column_config(),
    )


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

    st.dataframe(
        _format_table(filtered),
        width="stretch",
        height=560,
        column_config=_table_column_config(),
    )


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


def _table_column_config() -> dict[str, object]:
    numeric_columns = {
        "Продажи, шт.",
        "Остаток, шт.",
        "Продажи/день",
        "История/день",
        "Прогноз/день",
        "Горизонт, дней",
        "Спрос на горизонт, шт.",
        "Запас, дней",
        "Закупаем вперед, дней",
        "Страховой запас, шт.",
        "Нужно иметь, шт.",
        "Дефицит до цели, шт.",
        "Заказать, шт.",
        "Закупка/шт.",
        "Сумма закупки",
        "Выручка",
        "Маржа",
        "Ист. прогонов",
    }
    return {
        column: (
            st.column_config.NumberColumn(column, help=help_text)
            if column in numeric_columns
            else st.column_config.TextColumn(column, help=help_text)
        )
        for column, help_text in COLUMN_TOOLTIPS.items()
    }


def _empty_state(runs: pd.DataFrame) -> None:
    st.info(
        "Выберите источник данных и загрузите 4 Excel-файла: продажи USA, продажи WW, остатки USA и остатки WW. "
        "После загрузки расчет обновится автоматически."
    )


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
