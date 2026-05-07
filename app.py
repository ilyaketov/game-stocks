"""
Стоки игр — облачная версия (Streamlit Community Cloud).

Сессионный режим: загружаешь файлы → получаешь отчёты → скачиваешь.
Между сессиями ничего не сохраняется (это особенность Streamlit Cloud).

Локальная версия с историей живёт отдельно — см. README.
"""
from pathlib import Path
import sys

APP_DIR = Path(__file__).parent
sys.path.insert(0, str(APP_DIR))

import streamlit as st  # noqa: E402

from core import process_uploaded_files, build_report_for_month, report_filename_for_period  # noqa: E402


# ===== Конфигурация =====
st.set_page_config(
    page_title="Стоки игр",
    page_icon="🎮",
    layout="wide",
    initial_sidebar_state="collapsed",
)

MONTHS_RU = {
    1: 'январь', 2: 'февраль', 3: 'март', 4: 'апрель',
    5: 'май', 6: 'июнь', 7: 'июль', 8: 'август',
    9: 'сентябрь', 10: 'октябрь', 11: 'ноябрь', 12: 'декабрь',
}


def fmt_period(year, month):
    return f"{MONTHS_RU[month].capitalize()} {year}"


# ===== Заголовок =====
st.markdown(
    "<h1 style='color:#305496;margin-bottom:0;'>🎮 Стоки игр</h1>"
    "<p style='color:#595959;margin-top:0;'>"
    "Учёт остатков игровых ключей по выгрузкам QuickBooks Inventory Report</p>",
    unsafe_allow_html=True,
)

with st.expander("ℹ️ Как пользоваться", expanded=False):
    st.markdown("""
1. **Загрузите xlsx-выгрузку** — одно- или многомесячную, или несколько файлов сразу.
2. **Приложение автоматически** определит периоды в данных и построит отчёты на конец каждого месяца.
3. **Скачайте** нужные отчёты.

**Важно — это облачная версия:** данные не сохраняются между сессиями. Чтобы получить корректные «Дни с последнего закупа» для какого-то месяца, нужно загрузить **всю историю** — либо многомесячным файлом, либо несколькими одномесячными за все месяцы.
""")

st.divider()


# ===== Загрузка файлов =====
st.subheader("📥 Загрузить выгрузки")

uploaded = st.file_uploader(
    "Выберите один или несколько xlsx-файлов",
    type=['xlsx'],
    accept_multiple_files=True,
    help="Можно загружать одно- и многомесячные выгрузки одновременно.",
)

# Парсим только если набор файлов изменился (или впервые)
if uploaded:
    files_signature = tuple(sorted((f.name, len(f.getvalue())) for f in uploaded))
    if st.session_state.get('_last_signature') != files_signature:
        try:
            with st.spinner("Парсинг файлов..."):
                data = process_uploaded_files(uploaded)
            st.session_state['_data'] = data
            st.session_state['_last_signature'] = files_signature
        except ValueError as e:
            st.error(str(e))
            st.session_state.pop('_data', None)
            st.session_state.pop('_last_signature', None)
            st.stop()


data = st.session_state.get('_data')


# ===== Что прочитали =====
if data:
    with st.container(border=True):
        st.markdown("**📂 Прочитано из файлов:**")
        for f in data['file_summary']:
            periods_str = ", ".join(fmt_period(y, m) for y, m in f['periods'])
            badge = "📅" if len(f['periods']) == 1 else "📅📅"
            st.markdown(f"- {badge} `{f['filename']}` — {periods_str}")

        st.markdown(
            f"**Всего охвачено периодов:** {len(data['all_periods'])} "
            f"({fmt_period(*data['all_periods'][0])} — {fmt_period(*data['all_periods'][-1])})"
        )


# ===== Отчёты =====
st.divider()
st.subheader("📊 Отчёты")

if not data:
    st.info("👆 Загрузите файлы выше, чтобы построить отчёты.")
else:
    st.caption("Каждый отчёт строится на конец соответствующего месяца с учётом всей загруженной истории.")

    cols = st.columns(2)
    for i, (y, m) in enumerate(data['all_periods']):
        col = cols[i % 2]
        with col:
            with st.container(border=True):
                period_label = fmt_period(y, m)
                st.markdown(f"**📊 {period_label}**")

                # Кэшируем сгенерированный отчёт в session_state, чтобы повторное скачивание
                # не запускало генерацию заново.
                cache_key = f'_report_{y}_{m}_{st.session_state.get("_last_signature", "")}'

                if cache_key not in st.session_state:
                    with st.spinner(f"Строю отчёт за {period_label.lower()}..."):
                        try:
                            buf, flat = build_report_for_month(y, m, data)
                            st.session_state[cache_key] = {
                                'bytes': buf.getvalue(),
                                'rows': len(flat),
                                'end_qty': sum(r['end_qty'] for r in flat),
                                'end_val': sum(r['end_val'] for r in flat),
                            }
                        except Exception as e:
                            st.error(f"Ошибка: {e}")
                            continue

                rep = st.session_state[cache_key]
                st.caption(
                    f"Строк: {rep['rows']:,} • "
                    f"Остаток на конец: {rep['end_qty']:,.0f} шт. / ${rep['end_val']:,.2f}"
                )
                st.download_button(
                    label="⬇ Скачать xlsx",
                    data=rep['bytes'],
                    file_name=report_filename_for_period(y, m),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{y}_{m}",
                    use_container_width=True,
                )


# ===== Подвал =====
st.divider()
with st.expander("🛠 О приложении"):
    st.markdown("""
**Версия:** Облачная (Streamlit Community Cloud)

**Особенности:**
- Файлы и отчёты живут только в текущей сессии браузера. Закрытие вкладки → данные стираются.
- Поддерживаются и одномесячные (`Stock_2026-03.xlsx`), и многомесячные (`2025-12_-_2026-03.xlsx`) выгрузки.
- Для корректного расчёта «Дней с последнего закупа» приложение использует **всю** загруженную историю.

**Что в каждом отчёте:**
- **Источник** — копия исходной выгрузки за этот месяц
- **Свод** — итоги по подразделениям (на формулах SUMIF)
- **Таблица** — плоская таблица: одна строка = пара (артикул × магазин), с подсветкой по дням и жирным для позиций без движения
- **Отчёт** — то же, но сгруппировано по магазинам с подытогами
- **Риски** — анализ стареющего стока (3 категории), Топ-15 «мёртвых» позиций, выводы

**Если нужна история между сессиями** — есть локальная версия приложения, которая хранит файлы и отчёты в папке на компьютере.
""")
