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
    average_seasonality_factors,
    get_history_summary,
    get_run_items,
    list_runs,
    load_current_data,
    load_local_sample_data,
    MONTH_NAMES_RU,
    open_connection,
    read_seasonality_coefficients,
    resolve_local_sample_files,
    resolve_local_seasonality_file,
    save_analysis,
    summarize_result,
)


APP_DIR = Path(__file__).resolve().parent
DB_PATH = Path(os.getenv("INVENTORY_DB_PATH", APP_DIR / "data" / "inventory_history.sqlite"))
CRITICAL_COVER_DAYS = 14
DISPLAY_CURRENCY = "EUR"
DEFAULT_WW_SALES_WEBHOOK_URL = os.getenv(
    "WW_SALES_WEBHOOK_URL",
    "https://n8n.1703.team/webhook/mcp-paid-orders-report",
)

STATUS_LABELS = {
    "out_of_stock": "Нет остатка",
    "urgent": "Срочно",
    "planned": "Планово",
    "covered": "Закрыто",
    "no_demand": "Нет спроса",
}

METRIC_LABELS = {
    "revenue": f"Выручка, {DISPLAY_CURRENCY}",
    "margin": f"Маржа, {DISPLAY_CURRENCY}",
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
    "base_forecast_daily_sales": "Базовый прогноз/день",
    "seasonality_factor": "Коэф. сезонности",
    "forecast_daily_sales": "Прогноз/день",
    "purchase_horizon_days": "Горизонт, дней",
    "horizon_demand_qty": "Спрос на горизонт, шт.",
    "stock_cover_days": "Запас, дней",
    "target_cover_days": "Закупаем вперед, дней",
    "safety_stock_qty": "Страховой запас, шт.",
    "target_stock": "Нужно иметь, шт.",
    "raw_order_qty": "Дефицит до цели, шт.",
    "order_qty": "Заказать, шт.",
    "unit_purchase_cost": f"Закупка/шт., {DISPLAY_CURRENCY}",
    "order_cost_estimate": f"Сумма закупки, {DISPLAY_CURRENCY}",
    "revenue": f"Выручка, {DISPLAY_CURRENCY}",
    "margin": f"Маржа, {DISPLAY_CURRENCY}",
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
    "Базовый прогноз/день": "Дневной прогноз до сезонной корректировки: текущие продажи/день с учетом выбранного веса истории.",
    "Коэф. сезонности": "Средний коэффициент из листа сезонности по выбранным месяцам. 1.20 значит прогноз и закуп увеличены на 20%; 0.80 значит снижены на 20%.",
    "Прогноз/день": "Дневной спрос, который используется для закупа после сезонной корректировки. Формула: базовый прогноз/день * коэф. сезонности.",
    "Горизонт, дней": "На сколько дней вперед рассчитывается закуп.",
    "Спрос на горизонт, шт.": "Прогноз/день, умноженный на выбранный горизонт закупа.",
    "Запас, дней": "На сколько дней хватит текущего остатка: остаток / прогноз/день.",
    "Закупаем вперед, дней": "Целевой горизонт закупа из боковой панели.",
    "Страховой запас, шт.": "Дополнительный запас в штуках: прогноз/день * страховой запас в днях.",
    "Нужно иметь, шт.": "Сколько нужно иметь на складе под горизонт закупа и страховой запас.",
    "Дефицит до цели, шт.": "Нехватка до целевого остатка до округления по минимуму и кратности заказа.",
    "Заказать, шт.": "Итоговая рекомендация к заказу: дефицит после округления по минимальному заказу и кратности.",
    f"Закупка/шт., {DISPLAY_CURRENCY}": f"Расчетная закупочная цена за штуку из текущих продаж или истории. Валюта: {DISPLAY_CURRENCY}.",
    f"Сумма закупки, {DISPLAY_CURRENCY}": f"Оценка суммы закупки: заказать, шт. * закупка/шт. Валюта: {DISPLAY_CURRENCY}.",
    f"Выручка, {DISPLAY_CURRENCY}": f"Выручка из загруженного файла продаж за выбранный период. Валюта: {DISPLAY_CURRENCY}.",
    f"Маржа, {DISPLAY_CURRENCY}": f"Маржа из загруженного файла продаж за выбранный период. Валюта: {DISPLAY_CURRENCY}.",
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
    (
        config,
        run_name,
        notes,
        use_local_files,
        save_to_history,
        auto_sales_period,
        use_ww_sales_webhook,
        ww_sales_webhook_url,
    ) = _sidebar()
    input_sources = _file_inputs(use_local_files, use_ww_sales_webhook, ww_sales_webhook_url)
    seasonality_coefficients = _apply_seasonality_config(config, input_sources.get("seasonality"))
    _calculation_context(config, auto_sales_period)
    _seasonality_preview(seasonality_coefficients, config)

    runs = list_runs(conn)
    if "latest_result" not in st.session_state:
        st.session_state["latest_result"] = None
        st.session_state["latest_run_uuid"] = None

    required_sources = ["usa_sales", "usa_stock", "ww_stock"]
    if not use_ww_sales_webhook:
        required_sources.append("ww_sales")
    ready = all(input_sources.get(key) for key in required_sources)
    if use_ww_sales_webhook and not ww_sales_webhook_url.strip():
        ready = False
    latest_result = None
    current_run_uuid = None
    if ready:
        try:
            with st.spinner("Расчет..."):
                sales, stock = (
                    load_local_sample_data(config.period_start, config.period_end)
                    if use_local_files and not use_ww_sales_webhook
                    else load_current_data(
                        input_sources["usa_sales"],
                        input_sources.get("ww_sales"),
                        input_sources["usa_stock"],
                        input_sources["ww_stock"],
                        period_start=config.period_start,
                        period_end=config.period_end,
                        ww_sales_webhook_url=ww_sales_webhook_url if use_ww_sales_webhook else None,
                    )
                )
                history = get_history_summary(conn)
                latest_result = compute_analysis(sales, stock, history, config)
                st.session_state["latest_result"] = latest_result
                st.session_state["latest_run_uuid"] = None
        except Exception as exc:
            st.error(f"Не удалось прочитать файлы или выполнить расчет: {exc}")
            st.stop()

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
            3. Названия колонок могут немного отличаться: дашборд ищет синонимы вроде `Закупочная стоимость товаров`, `Средняя закупочная цена`, `Розничная цена`, `Доля от стоимости товаров`.
            4. Проверьте `Последняя дата продаж`. Например, если закупаем на 90 дней вперед и последняя дата 22.07.2026, дашборд считает продажи за 24.04.2026-22.07.2026.
            5. Расчет обновится автоматически после изменения параметра.
            6. Если результат нужно оставить в истории, нажмите `Сохранить текущий расчет в историю`.
            7. На вкладке `Закупка` смотрите позиции, которые нужно заказать. Фильтры сверху ограничивают рынок,
               ABCD-категории, статус остатка и поиск по товару или поставщику.
            8. На вкладке `Нет спроса` смотрите товары, которые есть в остатках, но не продавались в выбранном периоде.
            9. На вкладке `Экспорт` скачайте закупку, товары без спроса или все SKU в CSV/XLSX.

            **Как читать результат**

            - `A` - товары, которые дают основную долю выбранной метрики, по умолчанию выручки.
            - `B` и `C` - средний хвост продаж.
            - `D` - низкий вклад или нет текущего спроса.
            - `Нет спроса` - товар есть в текущих остатках, но продажи за выбранный период равны 0.
            - `Нет остатка` и `Срочно` - позиции с самым высоким приоритетом.
            - `Планово` - позицию стоит заказать до целевого уровня запаса.
            - `Закрыто` - текущего остатка достаточно.
            - Главное поле для закупа - `Заказать, шт.`. Статус только объясняет срочность.
            - `Наличие A по упаковкам` считает долю упаковок/SKU группы A, где остаток больше 0. Количество штук внутри упаковки в этот процент не входит.

            **Формула заказа**

            `Заказать, шт. = прогноз продаж в день * выбранный горизонт закупа + страховой запас в штуках - текущий остаток`.
            В режиме авто-периода `прогноз продаж в день` считается по продажам за такое же количество дней назад, как горизонт закупа вперед.
            Если файл продаж агрегированный и в нем нет дат по строкам, выгружайте файл продаж сразу за этот период.
            Если нет готовой маржинальности, она считается как `Маржа / Стоимость товаров`; если нет маржи, но есть закупочная стоимость, маржа считается как `Стоимость товаров - Закупочная стоимость`.
            История не создает закуп при нулевых текущих продажах, пока не включен отдельный чекбокс `Использовать историю при 0 продаж`.
            По умолчанию текущие продажи главнее истории. Историю можно включить отдельно для старых SKU или вручную поднять ее вес в прогнозе.
            """
        )


def _apply_seasonality_config(
    config: AnalysisConfig,
    seasonality_source: object | None,
) -> pd.DataFrame | None:
    config.seasonality_factors = {}
    if not config.seasonality_months:
        return None

    if seasonality_source is None:
        st.warning(
            "Сезонность включена, но файл сезонности не загружен. "
            "Коэффициент будет 1.00, закуп не изменится."
        )
        return None

    try:
        coefficients = read_seasonality_coefficients(seasonality_source)
        factors = average_seasonality_factors(coefficients, config.seasonality_months)
    except Exception as exc:
        st.warning(f"Не удалось прочитать сезонность: {exc}. Коэффициент будет 1.00.")
        return None

    if not factors:
        st.warning("Для выбранных месяцев не найдены коэффициенты сезонности. Коэффициент будет 1.00.")
        return coefficients

    config.seasonality_factors = factors
    st.sidebar.caption(
        "Коэффициент сезонности: "
        f"USA {factors.get('USA', factors.get('USA+WW', 1.0)):.3f}; "
        f"WW {factors.get('WW', factors.get('USA+WW', 1.0)):.3f}"
    )
    return coefficients


def _seasonality_preview(coefficients: pd.DataFrame | None, config: AnalysisConfig) -> None:
    if coefficients is None or not config.seasonality_months:
        return

    selected = coefficients[coefficients["month"].isin(config.seasonality_months)].copy()
    if selected.empty:
        return

    selected_display = selected.rename(
        columns={
            "month_name": "Месяц",
            "USA": "USA",
            "WW": "WW",
            "USA+WW": "USA+WW",
        }
    )[["Месяц", "USA", "WW", "USA+WW"]]
    factors = config.seasonality_factors or {}
    with st.expander("Коэффициенты сезонности", expanded=False):
        st.caption(
            "Для закупа используется блок `Продажи товаров`. "
            "Если выбрано несколько месяцев, дашборд берет среднее значение коэффициента по выбранным месяцам."
        )
        st.dataframe(selected_display, width="stretch", hide_index=True)
        st.caption(
            f"Итоговый средний коэффициент: USA {factors.get('USA', factors.get('USA+WW', 1.0)):.3f}; "
            f"WW {factors.get('WW', factors.get('USA+WW', 1.0)):.3f}."
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
    if config.seasonality_months:
        month_names = ", ".join(MONTH_NAMES_RU[month] for month in config.seasonality_months)
        usa_factor = config.seasonality_factors.get("USA", config.seasonality_factors.get("USA+WW", 1.0))
        ww_factor = config.seasonality_factors.get("WW", config.seasonality_factors.get("USA+WW", 1.0))
        seasonality_rule = (
            f"Сезонность включена для месяцев: **{month_names}**. "
            f"Средний коэффициент: USA **{usa_factor:.3f}**, WW **{ww_factor:.3f}**. "
            "`Прогноз/день = Базовый прогноз/день * Коэф. сезонности`. "
        )
    else:
        seasonality_rule = "Сезонность выключена: коэффициент сезонности равен 1.00. "

    st.info(
        f"Период продаж для расчета: **{period_start} - {period_end}** "
        f"(**{config.period_days} дней**). "
        f"Все финансовые показатели показываются в **{DISPLAY_CURRENCY}**. "
        "Строки продаж из Excel сначала фильтруются по этому периоду, затем группируются в SKU-строки "
        "`рынок + поставщик + товар + упаковка`; поэтому количество SKU обычно сильно меньше количества строк продаж или проданных штук. "
        f"{period_rule} "
        f"`Продажи/день = Продажи, шт. / {config.period_days}`. "
        f"`Спрос на горизонт = Прогноз/день * {config.purchase_horizon_days} дней`. "
        f"`Заказать, шт. = max(0, Спрос на горизонт + Страховой запас - Остаток)`, "
        "затем применяется минимальный заказ и кратность. "
        f"{seasonality_rule}"
        "Если файл продаж агрегированный без дат по строкам, его нужно выгружать за этот же период; "
        "иначе старые продажи из большого файла нельзя автоматически исключить. "
        f"{history_rule}"
    )


def _sidebar() -> tuple[AnalysisConfig, str, str, bool, bool, bool, bool, str]:
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

    st.sidebar.header("Сезонность")
    use_seasonality = st.sidebar.checkbox(
        "Учитывать сезонность",
        value=False,
        help="Если включено, прогноз/день умножается на средний коэффициент выбранных месяцев из файла сезонности.",
    )
    seasonality_months: list[int] = []
    if use_seasonality:
        st.sidebar.caption("Выберите месяцы сезона")
        month_columns = st.sidebar.columns(3)
        for month, month_name in MONTH_NAMES_RU.items():
            if month_columns[(month - 1) % 3].checkbox(
                month_name,
                value=False,
                key=f"seasonality_month_{month}",
            ):
                seasonality_months.append(month)
        if not seasonality_months:
            st.sidebar.warning("Выберите хотя бы один месяц, иначе коэффициент сезонности будет 1.00.")

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
    local_seasonality = resolve_local_seasonality_file()
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
    use_ww_sales_webhook = st.sidebar.checkbox(
        "WW продажи из n8n webhook",
        value=True,
        help="Если включено, файл продаж WW не нужен: дашборд отправляет POST-запрос в n8n и читает XLSX-ответ как продажи WW.",
    )
    ww_sales_webhook_url = st.sidebar.text_input(
        "Webhook WW продаж",
        value=DEFAULT_WW_SALES_WEBHOOK_URL,
        disabled=not use_ww_sales_webhook,
        help="URL n8n webhook. Дашборд отправляет JSON: market, period_start, period_end.",
    )
    if use_local_files:
        st.sidebar.caption(
            "Локальные файлы: "
            + "; ".join(f"{name}: {path.name}" for name, path in local_files.items())
        )
        if local_seasonality:
            st.sidebar.caption(f"Сезонность: {local_seasonality.name}")
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
        seasonality_months=seasonality_months,
    )
    return (
        config,
        run_name,
        notes,
        use_local_files,
        save_to_history,
        auto_sales_period,
        use_ww_sales_webhook,
        ww_sales_webhook_url,
    )


def _file_inputs(
    use_local_files: bool,
    use_ww_sales_webhook: bool,
    ww_sales_webhook_url: str,
) -> dict[str, object]:
    if use_local_files:
        inputs = {key: path for key, path in resolve_local_sample_files().items()}
        if use_ww_sales_webhook:
            inputs["ww_sales"] = None
            inputs["ww_sales_webhook_url"] = ww_sales_webhook_url
        inputs["seasonality"] = resolve_local_seasonality_file()
        return inputs

    st.subheader("Загрузка новых файлов")
    st.caption(
        "Загрузите Excel-файлы продаж/остатков. Если включен `WW продажи из n8n webhook`, файл продаж WW не нужен. "
        "Файл сезонности нужен только если включен чекбокс `Учитывать сезонность`. "
        "После загрузки расчет обновится автоматически. "
        "Если файл продаж без дат по строкам, выгружайте его сразу за период продаж, показанный выше."
    )
    cols = st.columns(5)
    with cols[0]:
        usa_sales = st.file_uploader(
            "Продажи USA",
            type=["xlsx"],
            key="usa_sales",
            help="Файл продаж USA за период, который указан слева в полях `Период продаж с/по`.",
        )
    with cols[1]:
        if use_ww_sales_webhook:
            ww_sales = None
            st.info("Продажи WW будут загружены из n8n webhook.")
        else:
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
    with cols[4]:
        seasonality = st.file_uploader(
            "Сезонность",
            type=["xlsx"],
            key="seasonality",
            help="Excel с листом `Итог сезонность` или `Сезонность`. Для закупа берется блок `Продажи товаров`.",
        )
    return {
        "usa_sales": usa_sales,
        "ww_sales": ww_sales,
        "usa_stock": usa_stock,
        "ww_stock": ww_stock,
        "seasonality": seasonality,
        "ww_sales_webhook_url": ww_sales_webhook_url if use_ww_sales_webhook else "",
    }


def _render_dashboard(result: pd.DataFrame, runs: pd.DataFrame, run_uuid: str | None) -> None:
    summary = summarize_result(result)
    _metrics(summary, run_uuid)
    _a_stock_availability_panel(result)

    purchase, filtered = _filters(result)
    no_demand = filtered[filtered["stock_status"].eq("no_demand")].copy()
    tabs = st.tabs(["Закупка", "Нет спроса", "ABCD", "История", "Экспорт"])

    with tabs[0]:
        _purchase_tab(purchase, filtered)
    with tabs[1]:
        _no_demand_tab(no_demand)
    with tabs[2]:
        _abcd_tab(filtered)
    with tabs[3]:
        _history_tab(runs)
    with tabs[4]:
        _export_tab(purchase, no_demand, result, summary)


def _metrics(summary: dict[str, float], run_uuid: str | None) -> None:
    cols = st.columns(8)
    cols[0].metric(
        "SKU всего",
        f"{summary['items_count']:,.0f}",
        help="Количество уникальных строк `рынок + поставщик + товар + упаковка`, попавших в текущий расчет. Это не строки продаж из Excel: продажи одной упаковки сворачиваются в одну SKU-строку.",
    )
    cols[1].metric(
        "Продано, шт.",
        f"{summary['total_sales_qty']:,.0f}",
        help="Сумма количества продаж из файлов за выбранный период после фильтра по датам. Это число сопоставимо с общими продажами из Excel, если выбран весь период.",
    )
    cols[2].metric(
        "К заказу",
        f"{summary['recommended_items_count']:,.0f}",
        help="Сколько SKU получили рекомендацию `Заказать, шт.` больше 0.",
    )
    cols[3].metric(
        "Заказать, шт.",
        f"{summary['total_order_qty']:,.0f}",
        help="Общее количество штук к заказу по всем отфильтрованным SKU.",
    )
    cols[4].metric(
        f"Сумма, {DISPLAY_CURRENCY}",
        f"{summary['total_order_cost_estimate']:,.0f}",
        help=f"Оценка стоимости закупа: сумма `Заказать, шт. * Закупка/шт.`. Валюта: {DISPLAY_CURRENCY}.",
    )
    cols[5].metric(
        f"Выручка, {DISPLAY_CURRENCY}",
        f"{summary['total_revenue']:,.0f}",
        help=f"Выручка из загруженных файлов продаж за выбранный период продаж. Валюта: {DISPLAY_CURRENCY}.",
    )
    cols[6].metric(
        "Срочно",
        f"{summary['urgent_items_count']:,.0f}",
        help="Количество SKU со статусом `Нет остатка` или `Срочно`.",
    )
    cols[7].metric(
        "Нет спроса",
        f"{summary['no_demand_items_count']:,.0f}",
        help="SKU, которые есть в остатках или данных, но продажи за выбранный период равны 0.",
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
        default=list(STATUS_LABELS.keys()),
        format_func=lambda value: STATUS_LABELS[value],
        help="Фильтр по статусу. По умолчанию включены все статусы, включая товары без спроса.",
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
                labels={
                    "order_qty": "Штук",
                    "order_cost_estimate": f"Сумма закупки, {DISPLAY_CURRENCY}",
                    "seedbank": "Поставщик",
                    "market": "Рынок",
                },
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


def _no_demand_tab(no_demand: pd.DataFrame) -> None:
    st.info(
        "`Нет спроса` - SKU, которые есть в остатках/текущих данных, но продажи за выбранный период равны 0. "
        "Они не попадают в закуп по умолчанию, но доступны для анализа и отдельного экспорта."
    )
    if no_demand.empty:
        st.success("По текущим фильтрам товаров без спроса нет.")
        return

    metric_cols = st.columns(3)
    metric_cols[0].metric(
        "SKU без спроса",
        f"{len(no_demand):,}",
        help="Количество SKU со статусом `Нет спроса` после выбранных фильтров.",
    )
    metric_cols[1].metric(
        "Остаток, шт.",
        f"{no_demand['current_stock'].sum():,.0f}",
        help="Суммарный остаток по SKU без продаж в выбранном периоде.",
    )
    metric_cols[2].metric(
        "Поставщиков",
        f"{no_demand['seedbank'].nunique():,.0f}",
        help="Количество поставщиков среди SKU без спроса.",
    )

    by_supplier = (
        no_demand.groupby(["market", "seedbank"], as_index=False)
        .agg(sku=("sku", "count"), current_stock=("current_stock", "sum"))
        .sort_values(["sku", "current_stock"], ascending=False)
        .head(30)
    )
    if not by_supplier.empty:
        fig = px.bar(
            by_supplier,
            x="sku",
            y="seedbank",
            color="market",
            orientation="h",
            labels={"sku": "SKU без спроса", "seedbank": "Поставщик", "market": "Рынок"},
            height=430,
        )
        fig.update_layout(yaxis={"categoryorder": "total ascending"}, margin=dict(l=10, r=10, t=20, b=10))
        st.plotly_chart(fig, width="stretch")

    st.dataframe(
        _format_table(no_demand.sort_values(["current_stock", "market"], ascending=[False, True])),
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
            labels={"abc_category": "ABCD", "revenue": f"Выручка, {DISPLAY_CURRENCY}", "market": "Рынок"},
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
            "total_revenue": f"Выручка, {DISPLAY_CURRENCY}",
            "total_margin": f"Маржа, {DISPLAY_CURRENCY}",
            "total_order_qty": "Заказать",
            "total_order_cost_estimate": f"Сумма закупки, {DISPLAY_CURRENCY}",
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
            labels={
                "created_at": "Дата",
                "value": "Значение",
                "variable": "Метрика",
                "total_order_cost_estimate": f"Сумма закупки, {DISPLAY_CURRENCY}",
            },
            height=420,
        )
        fig.update_layout(margin=dict(l=10, r=10, t=20, b=10))
        st.plotly_chart(fig, width="stretch")


def _export_tab(
    purchase: pd.DataFrame,
    no_demand: pd.DataFrame,
    result: pd.DataFrame,
    summary: dict[str, float],
) -> None:
    st.caption("Экспорт `Нет спроса` включает товары из остатков, у которых продажи за выбранный период равны 0.")
    export_cols = st.columns(3)
    with export_cols[0]:
        st.subheader("Закупка")
        st.download_button(
            "CSV закупка",
            dataframe_to_csv_bytes(_format_table(purchase)),
            file_name="purchase_recommendation.csv",
            mime="text/csv",
            disabled=purchase.empty,
        )
        st.download_button(
            "XLSX закупка",
            dataframe_to_xlsx_bytes(_format_table(purchase), summary),
            file_name="purchase_recommendation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            disabled=purchase.empty,
        )
    with export_cols[1]:
        st.subheader("Нет спроса")
        st.download_button(
            "CSV нет спроса",
            dataframe_to_csv_bytes(_format_table(no_demand)),
            file_name="no_demand_sku.csv",
            mime="text/csv",
            disabled=no_demand.empty,
        )
        st.download_button(
            "XLSX нет спроса",
            dataframe_to_xlsx_bytes(_format_table(no_demand), summary),
            file_name="no_demand_sku.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            disabled=no_demand.empty,
        )
    with export_cols[2]:
        st.subheader("Все SKU")
        st.download_button(
            "CSV все SKU",
            dataframe_to_csv_bytes(_format_table(result)),
            file_name="inventory_abcd_all_sku.csv",
            mime="text/csv",
        )
        st.download_button(
            "XLSX все SKU",
            dataframe_to_xlsx_bytes(_format_table(result), summary),
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
        "base_forecast_daily_sales",
        "seasonality_factor",
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
        "Базовый прогноз/день",
        "Коэф. сезонности",
        "Прогноз/день",
        "Горизонт, дней",
        "Спрос на горизонт, шт.",
        "Запас, дней",
        "Закупаем вперед, дней",
        "Страховой запас, шт.",
        "Нужно иметь, шт.",
        "Дефицит до цели, шт.",
        "Заказать, шт.",
        f"Закупка/шт., {DISPLAY_CURRENCY}",
        f"Сумма закупки, {DISPLAY_CURRENCY}",
        f"Выручка, {DISPLAY_CURRENCY}",
        f"Маржа, {DISPLAY_CURRENCY}",
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
