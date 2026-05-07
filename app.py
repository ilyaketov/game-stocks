"""
Веб-приложение KeyFlow — автоматизация отчётов закупа для QuickBooks.

Запуск:
    streamlit run app.py

Дизайн вдохновлён rokky.com:
- Фирменный акцент: индиго #4B4BFF
- Белый фон, минимализм, крупные числа
- Чистая типографика (Inter), большие отступы
- Прогрессивное раскрытие шагов (1 → 2 → 3)
"""
from __future__ import annotations

import hashlib
import io
import shutil
import tempfile
import time
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from engine import Pipeline
from config import PLOSHADKA_MAP


# ===========================================================================
# Helpers для инфографики
# ===========================================================================
# Палитра — оттенки фирменного индиго от насыщенного к светлому
TREEMAP_COLORS = ["#4B4BFF", "#7C5CFF", "#9F8DFF", "#B6A8FF", "#CFC5FF",
                  "#DAD3FF", "#E8E3FF", "#EFEAFF", "#F5F2FF"]


def _fmt_money(v: float) -> str:
    """Форматирует $1234567.89 → '$1 234 567.89'."""
    return f"${v:,.2f}".replace(",", " ")


def _fmt_int(v: int) -> str:
    """Форматирует 1234567 → '1 234 567'."""
    return f"{v:,}".replace(",", " ")


def _squarify(values: list[float], width: float, height: float) -> list[dict]:
    """Простой squarified treemap — рекурсивно режет область пропорционально значениям.

    Возвращает список {x, y, w, h} в том же порядке, что values.
    """
    if not values:
        return []
    total = sum(values)
    if total <= 0:
        return [{"x": 0, "y": 0, "w": 0, "h": 0} for _ in values]

    # Сортируем по убыванию, потом восстановим порядок
    indexed = sorted(enumerate(values), key=lambda x: -x[1])
    rects = [None] * len(values)

    def layout(items, x0, y0, w, h):
        if not items:
            return
        if len(items) == 1:
            i, _ = items[0]
            rects[i] = {"x": x0, "y": y0, "w": w, "h": h}
            return

        # Делим на две группы: первая ~ половина суммы
        total_local = sum(v for _, v in items)
        half = total_local / 2
        cumsum = 0
        split_idx = 1
        for k, (_, v) in enumerate(items):
            cumsum += v
            if cumsum >= half:
                split_idx = k + 1
                break

        first = items[:split_idx]
        rest = items[split_idx:]
        first_sum = sum(v for _, v in first)
        rest_sum = total_local - first_sum

        # Режем по длинной стороне
        if w >= h:
            split_w = w * (first_sum / total_local) if total_local else 0
            layout(first, x0, y0, split_w, h)
            layout(rest, x0 + split_w, y0, w - split_w, h)
        else:
            split_h = h * (first_sum / total_local) if total_local else 0
            layout(first, x0, y0, w, split_h)
            layout(rest, x0, y0 + split_h, w, h - split_h)

    layout(indexed, 0, 0, width, height)
    return rects


def _render_treemap(success: dict) -> None:
    """Концепт 1 — Treemap «Доли площадок»."""
    items = sorted(success.items(), key=lambda x: -x[1]["cost"])
    total_cost = sum(v["cost"] for _, v in items)
    total_qty = sum(v["qty"] for _, v in items)

    # KPI блок
    kpi_html = f"""
    <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 0.75rem; margin-bottom: 1rem;">
      <div style="background: white; padding: 1.1rem 1.25rem; border-radius: 12px; border: 1px solid #EDEDF2;">
        <div style="color: #6B6B80; font-size: 0.7rem; font-weight: 500; text-transform: uppercase; letter-spacing: 0.04em;">Себестоимость</div>
        <div style="color: #0A0A1F; font-size: 1.5rem; font-weight: 700; letter-spacing: -0.02em; margin-top: 0.4rem;">{_fmt_money(total_cost)}</div>
      </div>
      <div style="background: white; padding: 1.1rem 1.25rem; border-radius: 12px; border: 1px solid #EDEDF2;">
        <div style="color: #6B6B80; font-size: 0.7rem; font-weight: 500; text-transform: uppercase; letter-spacing: 0.04em;">Закуплено ключей</div>
        <div style="color: #0A0A1F; font-size: 1.5rem; font-weight: 700; letter-spacing: -0.02em; margin-top: 0.4rem;">{_fmt_int(total_qty)}</div>
        <div style="color: #6B6B80; font-size: 0.75rem; margin-top: 0.25rem;">в среднем {_fmt_money(total_cost/total_qty if total_qty else 0)} / ключ</div>
      </div>
      <div style="background: white; padding: 1.1rem 1.25rem; border-radius: 12px; border: 1px solid #EDEDF2;">
        <div style="color: #6B6B80; font-size: 0.7rem; font-weight: 500; text-transform: uppercase; letter-spacing: 0.04em;">Площадок</div>
        <div style="color: #0A0A1F; font-size: 1.5rem; font-weight: 700; letter-spacing: -0.02em; margin-top: 0.4rem;">{len(items)}</div>
      </div>
      <div style="background: white; padding: 1.1rem 1.25rem; border-radius: 12px; border: 1px solid #EDEDF2;">
        <div style="color: #6B6B80; font-size: 0.7rem; font-weight: 500; text-transform: uppercase; letter-spacing: 0.04em;">Уникальных SKU</div>
        <div style="color: #0A0A1F; font-size: 1.5rem; font-weight: 700; letter-spacing: -0.02em; margin-top: 0.4rem;">{sum(v['products'] for _, v in items)}</div>
      </div>
    </div>
    """
    st.markdown(kpi_html, unsafe_allow_html=True)

    # Treemap
    width, height = 660, 240
    values = [v["cost"] for _, v in items]
    rects = _squarify(values, width, height)

    svg_parts = [
        f'<svg viewBox="0 0 {width} {height}" style="width: 100%; height: auto; display: block;" '
        f'role="img" aria-label="Treemap по площадкам">'
    ]

    for idx, ((name, data), rect) in enumerate(zip(items, rects)):
        x, y, w, h = rect["x"], rect["y"], rect["w"], rect["h"]
        if w < 2 or h < 2:
            continue
        color = TREEMAP_COLORS[min(idx, len(TREEMAP_COLORS) - 1)]
        # Выбор цвета текста (тёмный на светлых блоках)
        is_dark_bg = idx <= 3
        text_color = "white" if is_dark_bg else "#26215C"
        text_color_sub = "rgba(255,255,255,0.85)" if is_dark_bg else "#534AB7"
        text_color_meta = "rgba(255,255,255,0.7)" if is_dark_bg else "#7F77DD"

        pct = data["cost"] / total_cost * 100 if total_cost else 0

        svg_parts.append(
            f'<rect x="{x+1}" y="{y+1}" width="{w-2}" height="{h-2}" rx="6" fill="{color}"/>'
        )

        # Название всегда влезает
        if w > 50 and h > 28:
            svg_parts.append(
                f'<text x="{x+12}" y="{y+22}" fill="{text_color}" '
                f'font-family="Inter,sans-serif" font-size="14" font-weight="700">{name}</text>'
            )
        # Сумма — если есть место
        if w > 70 and h > 50:
            svg_parts.append(
                f'<text x="{x+12}" y="{y+40}" fill="{text_color_sub}" '
                f'font-family="Inter,sans-serif" font-size="12" font-weight="500">'
                f'{_fmt_money(data["cost"])}</text>'
            )
        # Процент и ключи — если ещё больше места
        if w > 90 and h > 70:
            svg_parts.append(
                f'<text x="{x+12}" y="{y+57}" fill="{text_color_meta}" '
                f'font-family="Inter,sans-serif" font-size="11">'
                f'{pct:.1f}% · {_fmt_int(data["qty"])} ключей</text>'
            )

    svg_parts.append('</svg>')
    treemap_html = (
        f'<div style="background: white; border: 1px solid #EDEDF2; border-radius: 14px; padding: 1.25rem 1.4rem;">'
        f'<div style="display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 1rem;">'
        f'<div style="font-weight: 600; font-size: 0.95rem; color: #0A0A1F;">Доля площадок по себестоимости</div>'
        f'<div style="font-size: 0.75rem; color: #8B8B9E;">размер блока ∝ сумма закупа</div>'
        f'</div>{"".join(svg_parts)}</div>'
    )
    st.markdown(treemap_html, unsafe_allow_html=True)


def _render_breakdown_list(success: dict) -> None:
    """Концепт 4 — компактный список со всеми площадками + прогресс-бары."""
    items = sorted(success.items(), key=lambda x: -x[1]["cost"])
    total_cost = sum(v["cost"] for _, v in items)
    total_qty = sum(v["qty"] for _, v in items)

    rows_html = []
    for idx, (name, data) in enumerate(items):
        color = TREEMAP_COLORS[min(idx, len(TREEMAP_COLORS) - 1)]
        pct = data["cost"] / total_cost * 100 if total_cost else 0
        avg = data["cost"] / data["qty"] if data["qty"] else 0
        rows_html.append(f"""
        <div style="display: grid; grid-template-columns: 24px 110px 90px 1fr 130px 90px;
                    gap: 0.75rem; align-items: center; padding: 0.85rem 1.25rem;
                    border-bottom: 1px solid #F4F4F8;">
          <span style="display: inline-block; width: 8px; height: 8px; border-radius: 2px; background: {color};"></span>
          <span style="font-weight: 600; color: #0A0A1F; font-size: 0.9rem;">{name}</span>
          <span style="font-variant-numeric: tabular-nums; text-align: right; color: #0A0A1F; font-weight: 500; font-size: 0.85rem;">{_fmt_int(data['qty'])}</span>
          <div style="display: flex; align-items: center; gap: 0.5rem;">
            <div style="flex: 1; height: 6px; background: #F0F0F4; border-radius: 99px; overflow: hidden;">
              <div style="width: {pct:.1f}%; height: 100%; background: {color}; border-radius: 99px;"></div>
            </div>
            <span style="font-size: 0.75rem; color: #6B6B80; min-width: 40px; text-align: right; font-variant-numeric: tabular-nums;">{pct:.1f}%</span>
          </div>
          <span style="font-variant-numeric: tabular-nums; text-align: right; color: #0A0A1F; font-weight: 500; font-size: 0.85rem;">{_fmt_money(data['cost'])}</span>
          <span style="font-variant-numeric: tabular-nums; text-align: right; color: #6B6B80; font-size: 0.8rem;">{_fmt_money(avg)}</span>
        </div>""")

    avg_total = total_cost / total_qty if total_qty else 0
    total_row_html = f"""
    <div style="display: grid; grid-template-columns: 24px 110px 90px 1fr 130px 90px;
                gap: 0.75rem; align-items: center; padding: 0.85rem 1.25rem;
                background: #FAFAFB; border-top: 1px solid #EDEDF2;">
      <span></span>
      <span style="font-weight: 700; color: #0A0A1F; font-size: 0.9rem;">Итого</span>
      <span style="font-variant-numeric: tabular-nums; text-align: right; color: #0A0A1F; font-weight: 700; font-size: 0.9rem;">{_fmt_int(total_qty)}</span>
      <span></span>
      <span style="font-variant-numeric: tabular-nums; text-align: right; color: #0A0A1F; font-weight: 700; font-size: 0.9rem;">{_fmt_money(total_cost)}</span>
      <span style="font-variant-numeric: tabular-nums; text-align: right; color: #6B6B80; font-size: 0.8rem;">{_fmt_money(avg_total)}</span>
    </div>
    """

    header_html = """
    <div style="display: grid; grid-template-columns: 24px 110px 90px 1fr 130px 90px;
                gap: 0.75rem; align-items: center; padding: 0.6rem 1.25rem 0.5rem;
                font-size: 0.7rem; color: #8B8B9E; text-transform: uppercase;
                letter-spacing: 0.04em; font-weight: 500; border-bottom: 1px solid #EDEDF2;">
      <span></span>
      <span>Площадка</span>
      <span style="text-align: right;">Ключи</span>
      <span style="padding-left: 0.5rem;">Доля от всего</span>
      <span style="text-align: right;">Себестоимость</span>
      <span style="text-align: right;">Avg / ключ</span>
    </div>
    """

    list_html = (
        '<div style="background: white; border: 1px solid #EDEDF2; border-radius: 14px; overflow: hidden;">'
        + header_html + ''.join(rows_html) + total_row_html + '</div>'
    )
    st.markdown(list_html, unsafe_allow_html=True)


def _render_ploshadka_card(name: str, data: dict) -> None:
    """Концепт 2 — детальная карточка одной площадки."""
    ins = data["insights"]
    avg_price = ins["avg_price"]

    # Прогресс-бары топ-5 поставщиков
    suppliers_html = []
    for i, supp in enumerate(ins["top_suppliers"][:5]):
        bar_color = TREEMAP_COLORS[min(i, len(TREEMAP_COLORS) - 1)]
        # Ширина бара пропорциональна абсолютной доле, но минимум 2% для видимости
        bar_w = max(supp["pct_cost"], 1.5)
        suppliers_html.append(f"""
        <div style="margin-bottom: 0.85rem;">
          <div style="display: flex; justify-content: space-between; align-items: baseline; font-size: 0.85rem; margin-bottom: 0.3rem;">
            <span style="font-weight: 500; color: #0A0A1F;">{supp['name']}</span>
            <span style="color: #6B6B80; font-variant-numeric: tabular-nums; font-size: 0.8rem;">
              <span style="color: #0A0A1F; font-weight: 500;">{_fmt_money(supp['cost'])}</span>  ·  {supp['pct_cost']:.1f}%
            </span>
          </div>
          <div style="height: 6px; background: #F0F0F4; border-radius: 99px; overflow: hidden;">
            <div style="width: {bar_w:.1f}%; height: 100%; background: {bar_color}; border-radius: 99px;"></div>
          </div>
        </div>""")

    # Валютный сплит — горизонтальная stacked-полоска
    ccy_segments = []
    ccy_legend = []
    ccy_colors = {"USD": "#4B4BFF", "EUR": "#639922", "CNY": "#D85A30", "RUB": "#7F77DD"}
    for ccy in ins["ccy_split"]:
        col = ccy_colors.get(ccy["currency"], "#888780")
        w = max(ccy["pct_cost"], 0.3)  # минимум 0.3% чтобы было видно
        ccy_segments.append(
            f'<div style="width: {w:.2f}%; background: {col};"></div>'
        )
        ccy_legend.append(
            f'<span style="color: #4A4A5E;"><span style="color: {col}; font-weight: 600;">'
            f'{ccy["currency"]}</span> {_fmt_money(ccy["cost"])}</span>'
        )

    ccy_html = f"""
    <div>
      <div style="font-size: 0.7rem; color: #6B6B80; text-transform: uppercase; letter-spacing: 0.04em; font-weight: 500; margin-bottom: 0.4rem;">Валюты</div>
      <div style="display: flex; height: 8px; border-radius: 99px; overflow: hidden;">
        {''.join(ccy_segments)}
      </div>
      <div style="display: flex; gap: 0.85rem; margin-top: 0.5rem; font-size: 0.75rem; flex-wrap: wrap;">
        {''.join(ccy_legend)}
      </div>
    </div>
    """

    # Концентрация
    conc_html = f"""
    <div>
      <div style="font-size: 0.7rem; color: #6B6B80; text-transform: uppercase; letter-spacing: 0.04em; font-weight: 500; margin-bottom: 0.4rem;">Концентрация</div>
      <div style="font-size: 1.05rem; font-weight: 600; color: #0A0A1F;">{ins['concentration_pct']:.1f}%</div>
      <div style="font-size: 0.75rem; color: #6B6B80; margin-top: 0.15rem;">в одном поставщике ({ins['concentration_supplier']})</div>
    </div>
    """

    # Топ продукт
    top_prod = ins["top_product"]
    prod_name = (top_prod["name"][:35] + "…") if len(top_prod["name"]) > 35 else top_prod["name"]
    top_prod_html = f"""
    <div>
      <div style="font-size: 0.7rem; color: #6B6B80; text-transform: uppercase; letter-spacing: 0.04em; font-weight: 500; margin-bottom: 0.4rem;">Топ-продукт</div>
      <div style="font-size: 0.85rem; font-weight: 500; color: #0A0A1F; line-height: 1.3;">{prod_name}</div>
      <div style="font-size: 0.75rem; color: #6B6B80; margin-top: 0.15rem;">{_fmt_int(top_prod['qty'])} шт · {_fmt_money(top_prod['cost'])}</div>
    </div>
    """

    card_html = f"""
    <div style="background: white; border: 1px solid #EDEDF2; border-radius: 14px; padding: 1.5rem 1.75rem;">
      <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 1.5rem;">
        <div>
          <div style="display: flex; align-items: center; gap: 0.5rem; margin-bottom: 0.4rem;">
            <span style="display: inline-block; width: 10px; height: 10px; border-radius: 50%; background: #4B4BFF;"></span>
            <span style="font-weight: 700; font-size: 1.15rem; color: #0A0A1F;">{name}</span>
            <span style="background: #E8F5EE; color: #0F7A3E; padding: 0.15rem 0.55rem; border-radius: 999px; font-size: 0.7rem; font-weight: 500;">готов</span>
          </div>
          <div style="font-size: 0.8rem; color: #8B8B9E;">{_fmt_int(ins['total_qty'])} ключей · {ins['n_suppliers']} поставщиков · {ins['n_products']} SKU</div>
        </div>
        <div style="text-align: right;">
          <div style="font-size: 1.6rem; font-weight: 700; color: #0A0A1F; letter-spacing: -0.02em;">{_fmt_money(ins['total_cost'])}</div>
          <div style="font-size: 0.75rem; color: #6B6B80;">средняя цена {_fmt_money(avg_price)}</div>
        </div>
      </div>

      <div style="margin: 1.5rem 0 0.5rem;">
        <div style="display: flex; justify-content: space-between; font-size: 0.75rem; color: #6B6B80; margin-bottom: 0.6rem; font-weight: 500;">
          <span>топ-5 поставщиков</span>
          <span>доля себестоимости</span>
        </div>
        {''.join(suppliers_html)}
      </div>

      <div style="border-top: 1px solid #EDEDF2; margin-top: 1.25rem; padding-top: 1.25rem;
                  display: grid; grid-template-columns: 1.4fr 1fr 1fr; gap: 1.25rem;">
        {ccy_html}
        {conc_html}
        {top_prod_html}
      </div>
    </div>
    """
    st.markdown(card_html, unsafe_allow_html=True)


# ===========================================================================
# Настройка страницы
# ===========================================================================
st.set_page_config(
    page_title="KeyFlow",
    page_icon="◆",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ===========================================================================
# Стили — Rokky-inspired
# ===========================================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

/* === Глобальные === */
html, body, [class*="st-"], .stApp {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
}

.stApp {
    background: #FAFAFB;
}

/* Сужаем основной контейнер */
.main .block-container {
    max-width: 1100px;
    padding-top: 2.5rem;
    padding-bottom: 4rem;
}

/* === Прячем стандартный header === */
header[data-testid="stHeader"] {
    background: transparent;
    height: 0;
}

#MainMenu, footer {
    visibility: hidden;
}

/* === Заголовки === */
h1 {
    font-weight: 800 !important;
    font-size: 3rem !important;
    letter-spacing: -0.03em !important;
    color: #0A0A1F !important;
    line-height: 1.1 !important;
    margin-bottom: 0.5rem !important;
}

h2 {
    font-weight: 700 !important;
    font-size: 1.5rem !important;
    color: #0A0A1F !important;
    letter-spacing: -0.01em !important;
    margin-top: 2.5rem !important;
    margin-bottom: 1.25rem !important;
}

h3 {
    font-weight: 600 !important;
    font-size: 1rem !important;
    color: #0A0A1F !important;
    margin-bottom: 0.75rem !important;
}

p, .stMarkdown p {
    color: #4A4A5E;
    font-size: 0.95rem;
    line-height: 1.6;
}

/* Hero subtitle */
.hero-subtitle {
    font-size: 1.15rem;
    color: #6B6B80;
    font-weight: 400;
    margin-top: 0.25rem;
    margin-bottom: 2.5rem;
    max-width: 640px;
}

/* Step number */
.step-num {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 28px;
    height: 28px;
    background: #4B4BFF;
    color: white;
    border-radius: 50%;
    font-size: 0.85rem;
    font-weight: 600;
    margin-right: 0.6rem;
    vertical-align: middle;
}

.step-row {
    display: flex;
    align-items: center;
    margin-top: 2.5rem;
    margin-bottom: 1.25rem;
}

.step-row h2 {
    margin: 0 !important;
}

/* === Кнопки === */
.stButton > button, .stDownloadButton > button {
    font-family: 'Inter', sans-serif !important;
    font-weight: 500 !important;
    font-size: 0.9rem !important;
    border-radius: 8px !important;
    padding: 0.6rem 1.4rem !important;
    transition: all 0.15s ease !important;
    border: 1px solid transparent !important;
}

/* Primary кнопка — фирменный индиго */
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
    background: #4B4BFF !important;
    color: white !important;
    border: none !important;
    box-shadow: 0 1px 2px rgba(75, 75, 255, 0.15) !important;
}

.stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover {
    background: #3838E5 !important;
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(75, 75, 255, 0.25) !important;
}

/* Secondary кнопка */
.stButton > button[kind="secondary"], .stDownloadButton > button[kind="secondary"] {
    background: white !important;
    color: #0A0A1F !important;
    border: 1px solid #E4E4EA !important;
}

.stButton > button[kind="secondary"]:hover, .stDownloadButton > button[kind="secondary"]:hover {
    border-color: #4B4BFF !important;
    color: #4B4BFF !important;
}

/* === Загрузка файлов — крупные дроп-зоны === */
[data-testid="stFileUploader"] {
    background: white;
    border: 1.5px dashed #E4E4EA;
    border-radius: 14px;
    padding: 1rem;
    transition: all 0.15s ease;
}

[data-testid="stFileUploader"]:hover {
    border-color: #4B4BFF;
    background: #F8F8FF;
}

[data-testid="stFileUploader"] section {
    border: none !important;
    background: transparent !important;
    padding: 0.5rem !important;
}

[data-testid="stFileUploader"] section button {
    background: #F4F4F8 !important;
    color: #0A0A1F !important;
    border: none !important;
    font-size: 0.85rem !important;
}

/* === Метрики (Streamlit metric) === */
[data-testid="stMetric"] {
    background: white;
    padding: 1.25rem 1.5rem;
    border-radius: 12px;
    border: 1px solid #EDEDF2;
}

[data-testid="stMetric"] [data-testid="stMetricLabel"] {
    color: #6B6B80;
    font-size: 0.75rem !important;
    font-weight: 500 !important;
    text-transform: uppercase;
    letter-spacing: 0.04em;
}

[data-testid="stMetric"] [data-testid="stMetricValue"] {
    color: #0A0A1F;
    font-size: 1.75rem !important;
    font-weight: 700 !important;
    letter-spacing: -0.02em;
}

/* === Карточки-обёртки для блоков === */
.kf-card {
    background: white;
    border: 1px solid #EDEDF2;
    border-radius: 14px;
    padding: 1.5rem 1.75rem;
    margin-bottom: 1rem;
}

.kf-card-uploader {
    background: white;
    border: 1px solid #EDEDF2;
    border-radius: 14px;
    padding: 1.25rem;
    height: 100%;
}

.kf-uploader-label {
    font-weight: 600;
    font-size: 0.95rem;
    color: #0A0A1F;
    margin-bottom: 0.25rem;
}

.kf-uploader-hint {
    font-size: 0.8rem;
    color: #8B8B9E;
    margin-bottom: 1rem;
}

/* === Алёрты === */
[data-testid="stAlert"] {
    border-radius: 10px;
    border: none;
    padding: 0.85rem 1.1rem;
    font-size: 0.9rem;
}

div[data-baseweb="notification"][kind="info"] {
    background: #EEF0FF;
}

/* === Multiselect === */
[data-baseweb="select"] {
    border-radius: 10px !important;
}

[data-baseweb="tag"] {
    background: #EEF0FF !important;
    color: #4B4BFF !important;
    border-radius: 6px !important;
    font-weight: 500 !important;
}

/* === Таблицы === */
[data-testid="stDataFrame"] {
    border-radius: 12px;
    border: 1px solid #EDEDF2;
    overflow: hidden;
}

/* === Прогресс === */
.stProgress > div > div > div > div {
    background: #4B4BFF !important;
}

/* === Expander === */
[data-testid="stExpander"] {
    background: white;
    border: 1px solid #EDEDF2 !important;
    border-radius: 10px !important;
}

[data-testid="stExpander"] summary {
    font-weight: 500;
    font-size: 0.9rem;
}

/* === Caption === */
.caption-muted {
    color: #8B8B9E;
    font-size: 0.85rem;
    margin-top: 0.5rem;
}

/* Status pill */
.kf-pill {
    display: inline-block;
    padding: 0.2rem 0.7rem;
    background: #E8F5EE;
    color: #0F7A3E;
    font-size: 0.75rem;
    font-weight: 500;
    border-radius: 999px;
    letter-spacing: 0.01em;
}

.kf-pill-warn {
    background: #FFF4E5;
    color: #B25E00;
}

/* Логотип */
.kf-brand {
    display: flex;
    align-items: center;
    gap: 0.6rem;
    margin-bottom: 1rem;
}

.kf-brand-mark {
    width: 32px;
    height: 32px;
    background: linear-gradient(135deg, #4B4BFF 0%, #7C5CFF 100%);
    border-radius: 8px;
    display: flex;
    align-items: center;
    justify-content: center;
    color: white;
    font-weight: 700;
    font-size: 1rem;
    box-shadow: 0 2px 8px rgba(75, 75, 255, 0.25);
}

.kf-brand-name {
    font-size: 0.95rem;
    font-weight: 600;
    color: #0A0A1F;
    letter-spacing: -0.01em;
}

.kf-brand-sub {
    font-size: 0.85rem;
    color: #8B8B9E;
}

/* Discrete вертикальный отступ */
.kf-spacer {
    height: 1.5rem;
}

</style>
""", unsafe_allow_html=True)


# ===========================================================================
# Шапка / Hero
# ===========================================================================
st.markdown("""
<div class="kf-brand">
  <div class="kf-brand-mark">◆</div>
  <div>
    <div class="kf-brand-name">KeyFlow</div>
    <div class="kf-brand-sub">Purchase reports automation</div>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("# Отчёты закупа,<br>автоматически.", unsafe_allow_html=True)
st.markdown(
    '<p class="hero-subtitle">Загрузите выгрузки из биллинга — '
    'KeyFlow соберёт сводные отчёты для всех площадок в формате QuickBooks. '
    'Минуты вместо часов ручной работы.</p>',
    unsafe_allow_html=True,
)


# ===========================================================================
# Временная директория и кэш
# ===========================================================================
@st.cache_resource
def _make_tmpdir() -> Path:
    return Path(tempfile.mkdtemp(prefix="keyflow_"))


tmpdir = _make_tmpdir()


def _save_upload(uploaded_file, name: str) -> tuple[Path | None, str | None]:
    if uploaded_file is None:
        return None, None
    data = uploaded_file.getbuffer()
    digest = hashlib.sha1(bytes(data)).hexdigest()[:16]
    target = tmpdir / f"{digest}_{name}"
    if not target.exists():
        target.write_bytes(data)
    return target, digest


@st.cache_resource(show_spinner=False)
def _build_pipeline(r1_path: str | None, r2_path: str | None,
                    genba_path: str | None, _cache_key: str) -> Pipeline:
    return Pipeline(r1_path, r2_path, genba_path)


# ===========================================================================
# Шаг 1 — загрузка
# ===========================================================================
st.markdown(
    '<div class="step-row"><span class="step-num">1</span>'
    '<h2>Загрузите отчёты</h2></div>',
    unsafe_allow_html=True,
)

c1, c2, c3 = st.columns(3, gap="medium")

with c1:
    st.markdown('<div class="kf-uploader-label">Универсальный отчёт</div>', unsafe_allow_html=True)
    st.markdown('<div class="kf-uploader-hint">Старый формат биллинга</div>', unsafe_allow_html=True)
    file_r1 = st.file_uploader(
        "r1", type=["xlsx"], key="r1", label_visibility="collapsed",
    )

with c2:
    st.markdown('<div class="kf-uploader-label">Universal Report shipped</div>', unsafe_allow_html=True)
    st.markdown('<div class="kf-uploader-hint">Новый формат, ~310 тыс. строк</div>', unsafe_allow_html=True)
    file_r2 = st.file_uploader(
        "r2", type=["xlsx"], key="r2", label_visibility="collapsed",
    )

with c3:
    st.markdown('<div class="kf-uploader-label">genbaFile</div>', unsafe_allow_html=True)
    st.markdown('<div class="kf-uploader-hint">Цены закупа Genba</div>', unsafe_allow_html=True)
    file_genba = st.file_uploader(
        "genba", type=["xlsx"], key="genba", label_visibility="collapsed",
    )

if not (file_r1 or file_r2):
    st.markdown('<div class="kf-spacer"></div>', unsafe_allow_html=True)
    st.info("Загрузите хотя бы один из универсальных отчётов, чтобы продолжить.")
    st.stop()


p1_path, h1 = _save_upload(file_r1, "report1.xlsx")
p2_path, h2 = _save_upload(file_r2, "report2.xlsx")
pg_path, hg = _save_upload(file_genba, "genba.xlsx")
cache_key = f"{h1}|{h2}|{hg}"


# ===========================================================================
# Шаг 2 — анализ
# ===========================================================================
st.markdown(
    '<div class="step-row"><span class="step-num">2</span>'
    '<h2>Анализ</h2></div>',
    unsafe_allow_html=True,
)

with st.spinner("Парсим файлы (30–60 секунд при первой загрузке, дальше — мгновенно)..."):
    t0 = time.time()
    pipeline = _build_pipeline(
        str(p1_path) if p1_path else None,
        str(p2_path) if p2_path else None,
        str(pg_path) if pg_path else None,
        cache_key,
    )
    load_time = time.time() - t0

with st.spinner("Проверяем поставщиков..."):
    validation = pipeline.validate()

if not validation.available_ploshadki:
    st.error("В загруженных файлах не найдено ни одной известной площадки.")
    st.stop()

# Pill со статусом
n_ok = "Все распознаны" if not validation.unmapped_suppliers else f"{len(validation.unmapped_suppliers)} требует внимания"
pill_class = "kf-pill" if not validation.unmapped_suppliers else "kf-pill kf-pill-warn"
st.markdown(
    f'<span class="{pill_class}">{n_ok}</span> '
    f'<span class="caption-muted">·  загрузка за {load_time:.1f} с  ·  '
    f'{len(validation.available_ploshadki)} площадок найдено</span>',
    unsafe_allow_html=True,
)

st.markdown('<div class="kf-spacer"></div>', unsafe_allow_html=True)

# Метрики площадок — компактная сетка
n_ploshadki = len(validation.available_ploshadki)
n_per_row = min(n_ploshadki, 5)
items = list(validation.available_ploshadki.items())

for i in range(0, n_ploshadki, n_per_row):
    cols = st.columns(n_per_row, gap="small")
    chunk = items[i:i + n_per_row]
    for col, (key, n) in zip(cols, chunk):
        with col:
            st.metric(key, f"{n:,}".replace(",", " "))

# Неизвестные поставщики
if validation.unmapped_suppliers:
    st.markdown('<div class="kf-spacer"></div>', unsafe_allow_html=True)
    with st.expander(
        f"Неизвестные поставщики ({len(validation.unmapped_suppliers)}) — будут пропущены",
        expanded=False,
    ):
        st.caption(
            "Эти имена не сопоставлены с группами в эталоне. "
            "Чтобы строки попали в отчёт, добавьте маппинги в `SUPPLIER_MAPPING` (config.py)."
        )
        df_unmapped = pd.DataFrame([
            {"Имя в биллинге": k, "Строк": v}
            for k, v in sorted(validation.unmapped_suppliers.items(),
                               key=lambda x: -x[1])
        ])
        st.dataframe(df_unmapped, use_container_width=True, hide_index=True)


# ===========================================================================
# Шаг 3 — сборка
# ===========================================================================
st.markdown(
    '<div class="step-row"><span class="step-num">3</span>'
    '<h2>Соберите отчёты</h2></div>',
    unsafe_allow_html=True,
)

selected = st.multiselect(
    "Площадки",
    options=list(validation.available_ploshadki.keys()),
    default=list(validation.available_ploshadki.keys()),
    label_visibility="collapsed",
    placeholder="Выберите площадки",
)

if not selected:
    st.info("Выберите хотя бы одну площадку.")
    st.stop()

st.markdown('<div class="kf-spacer"></div>', unsafe_allow_html=True)

run_clicked = st.button("Собрать отчёты", type="primary", use_container_width=False)

if run_clicked:
    out_dir = tmpdir / f"out_{cache_key[:16]}"
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir()

    progress = st.progress(0, text="Запуск...")
    results = {}
    for i, key in enumerate(selected):
        progress.progress(i / len(selected), text=f"{key}...")
        try:
            t0 = time.time()
            agg = pipeline.aggregate(key)
            if agg.empty:
                results[key] = {"error": "Нет данных"}
                continue
            path = pipeline.save_to_excel(
                agg, key, out_dir / f"{key}_zakup_svod.xlsx"
            )
            insights = pipeline.insights(agg)
            results[key] = {
                "path": path,
                "qty": insights["total_qty"],
                "cost": insights["total_cost"],
                "suppliers": insights["n_suppliers"],
                "products": insights["n_products"],
                "elapsed": time.time() - t0,
                "insights": insights,
            }
        except Exception as e:
            results[key] = {"error": str(e)}

    progress.empty()
    # Сохраняем результаты в session_state, чтобы пережить перерисовку при кликах
    st.session_state["results"] = results
    st.session_state["out_dir"] = str(out_dir)

# Если результаты собраны (или ранее, или только что) — показываем
if "results" in st.session_state:
    results = st.session_state["results"]
    success = {k: v for k, v in results.items() if "error" not in v}
    failed = {k: v for k, v in results.items() if "error" in v}

    if success:
        # ===================================================================
        # Инфографика 1 — Treemap "Доли площадок"
        # ===================================================================
        st.markdown('<h3 style="margin-top: 1.5rem; margin-bottom: 1rem;">Сводка</h3>',
                    unsafe_allow_html=True)
        _render_treemap(success)

        st.markdown('<div class="kf-spacer"></div>', unsafe_allow_html=True)

        # ===================================================================
        # Инфографика 4 — Компактный список (главный экран)
        # ===================================================================
        _render_breakdown_list(success)

        st.markdown('<div class="kf-spacer"></div>', unsafe_allow_html=True)

        # ===================================================================
        # Инфографика 2 — Карточка площадки (по выбору)
        # ===================================================================
        st.markdown(
            '<h3 style="margin-top: 2rem; margin-bottom: 0.75rem;">Детали площадки</h3>',
            unsafe_allow_html=True,
        )
        ploshadka_for_detail = st.selectbox(
            "Площадка",
            options=list(success.keys()),
            label_visibility="collapsed",
            key="ploshadka_detail_select",
        )
        if ploshadka_for_detail:
            _render_ploshadka_card(ploshadka_for_detail, success[ploshadka_for_detail])

        st.markdown('<div class="kf-spacer"></div>', unsafe_allow_html=True)

        # ===================================================================
        # Скачивание
        # ===================================================================
        st.markdown(
            '<h3 style="margin-top: 2rem; margin-bottom: 0.75rem;">Скачать</h3>',
            unsafe_allow_html=True,
        )

        # ZIP со всеми
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for k, v in success.items():
                zf.write(v["path"], arcname=Path(v["path"]).name)
        zip_buffer.seek(0)

        dl_cols = st.columns([1, 3])
        with dl_cols[0]:
            st.download_button(
                "Скачать всё (ZIP)",
                data=zip_buffer,
                file_name="keyflow_reports.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True,
            )

        st.markdown('<div class="kf-spacer"></div>', unsafe_allow_html=True)
        st.markdown(
            '<p style="font-size: 0.85rem; color: #8B8B9E; margin-bottom: 0.5rem;">'
            'Или по отдельности:</p>',
            unsafe_allow_html=True,
        )

        files_per_row = 4
        items_l = list(success.items())
        for i in range(0, len(items_l), files_per_row):
            row_cols = st.columns(files_per_row, gap="small")
            chunk = items_l[i:i + files_per_row]
            for col, (k, v) in zip(row_cols, chunk):
                with col:
                    with open(v["path"], "rb") as f:
                        st.download_button(
                            k,
                            data=f.read(),
                            file_name=Path(v["path"]).name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_{k}",
                            use_container_width=True,
                        )

    if failed:
        st.markdown('<div class="kf-spacer"></div>', unsafe_allow_html=True)
        st.error("Не удалось обработать:")
        for k, v in failed.items():
            st.write(f"**{k}** — {v['error']}")


# ===========================================================================
# Подвал
# ===========================================================================
st.markdown('<div style="height: 4rem;"></div>', unsafe_allow_html=True)
st.markdown(
    '<div style="text-align: center; color: #B0B0C0; font-size: 0.8rem;">'
    'KeyFlow · Purchase reports automation</div>',
    unsafe_allow_html=True,
)
