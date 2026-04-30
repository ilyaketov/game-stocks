"""Расчёт плоской таблицы остатков за указанный календарный период."""
from datetime import datetime
from calendar import monthrange


def month_period(year, month):
    """Возвращает (period_start, period_end) для календарного месяца."""
    period_start = datetime(year, month, 1)
    last_day = monthrange(year, month)[1]
    period_end = datetime(year, month, last_day)
    return period_start, period_end


def build_flat(transactions, sku_to_game, period_start, period_end):
    """
    Строит плоскую таблицу (по парам артикул × магазин) на конец указанного периода.

    Логика остатка на конец: end_qty = SUM(Qty), end_val = SUM(Inventory cost)
    по всем транзакциям пары до period_end включительно. Это корректно работает
    даже на повреждённых выгрузках QuickBooks, где колонка Qty on hand может
    содержать ошибки отображения.

    Дни с последнего закупа считаются от последней Bill-транзакции для артикула
    в любом магазине (по всей доступной истории до period_end).
    """
    # Last bill by SKU (вся история до period_end)
    last_bill_by_sku = {}
    for (sku, store), txs in transactions.items():
        for dt, ttype, qty, cost, on_hand, asset in txs:
            if dt <= period_end and ttype == 'Bill':
                if sku not in last_bill_by_sku or dt > last_bill_by_sku[sku]:
                    last_bill_by_sku[sku] = dt

    flat = []
    for (sku, store), txs in transactions.items():
        txs_until_end = [t for t in txs if t[0] <= period_end]
        if not txs_until_end:
            continue

        end_qty = sum(t[2] for t in txs_until_end)
        end_val = sum(t[3] for t in txs_until_end)

        txs_in_period = [t for t in txs_until_end if t[0] >= period_start]

        bills_qty = sum(t[2] for t in txs_in_period if t[1] == 'Bill')
        bills_val = sum(t[3] for t in txs_in_period if t[1] == 'Bill')
        invoices_qty = sum(-t[2] for t in txs_in_period if t[1] == 'Invoice')
        invoices_val = sum(-t[3] for t in txs_in_period if t[1] == 'Invoice')
        adj_in_qty = sum(t[2] for t in txs_in_period if t[1] == 'Inventory Qty Adjust' and t[2] > 0)
        adj_in_val = sum(t[3] for t in txs_in_period if t[1] == 'Inventory Qty Adjust' and t[3] > 0)
        adj_out_qty = sum(-t[2] for t in txs_in_period if t[1] == 'Inventory Qty Adjust' and t[2] < 0)
        adj_out_val = sum(-t[3] for t in txs_in_period if t[1] == 'Inventory Qty Adjust' and t[3] < 0)

        start_qty = end_qty - bills_qty + invoices_qty - adj_in_qty + adj_out_qty
        start_val = end_val - bills_val + invoices_val - adj_in_val + adj_out_val

        days_since = None
        if sku in last_bill_by_sku:
            days_since = (period_end - last_bill_by_sku[sku]).days

        has_movement = len(txs_in_period) > 0
        has_any_sale_ever = any(t[1] == 'Invoice' for t in txs_until_end)

        flat.append({
            'store': store,
            'game': sku_to_game.get(sku, ''),
            'sku': sku,
            'upload_date': period_end,
            'start_qty': start_qty,
            'sold_qty': invoices_qty,
            'bought_qty': bills_qty,
            'end_qty': end_qty,
            'start_val': round(start_val, 2),
            'sold_val': round(invoices_val, 2),
            'bought_val': round(bills_val, 2),
            'end_val': round(end_val, 2),
            'days_since_bill': days_since,
            'has_movement_in_period': has_movement,
            'has_any_sale_ever': has_any_sale_ever,
        })

    return flat


def merge_transactions(*transaction_dicts):
    """
    Объединяет несколько словарей транзакций в один.

    ВАЖНО: дедупликация работает на уровне ПАР (sku, store), а не отдельных транзакций.
    Внутри одного файла транзакции уникальны и могут содержать «повторы»
    (например, две Invoice с одинаковым qty=-1 в один день — это две разные продажи).
    Поэтому объединение работает так:
      - Для каждой пары (sku, store) выбирается список транзакций из источника
        с НАИБОЛЬШИМ числом транзакций (как правило — из многомесячного файла).
      - Если разные источники дают одинаковую длину списка, остаются транзакции
        из первого встреченного.
      - Если один источник содержит транзакции, которых нет в другом
        (например, декабрьские в декабрьском файле + январские в январском),
        то списки объединяются и сортируются по дате.

    Это покрывает все встречающиеся сценарии:
      - 4 одномесячных файла (диапазоны не пересекаются → объединяем)
      - 1 многомесячный файл (один источник)
      - многомесячный + одномесячный с пересекающимися месяцами
        (берём многомесячный, т.к. в нём больше транзакций)
    """
    # Собираем по парам списки транзакций из всех источников
    by_pair = {}  # key -> list of (source_index, transactions_list)
    for idx, txs_dict in enumerate(transaction_dicts):
        for key, txs in txs_dict.items():
            by_pair.setdefault(key, []).append((idx, list(txs)))

    merged = {}
    for key, sources in by_pair.items():
        if len(sources) == 1:
            merged[key] = sorted(sources[0][1], key=lambda x: x[0])
            continue

        # Несколько источников. Проверим: если один из них «надмножество» другого
        # (по парам (date, type, qty, cost) с учётом порядка), берём его целиком.
        # Иначе объединяем диапазоны (по непересекающимся месяцам).

        # Вариант 1: непересекающиеся месяцы — просто склеиваем
        all_months_per_source = []
        for _, txs in sources:
            months = {(t[0].year, t[0].month) for t in txs}
            all_months_per_source.append(months)

        # Если все источники имеют непересекающиеся множества месяцев — объединяем как есть
        no_overlap = True
        seen = set()
        for ms in all_months_per_source:
            if ms & seen:
                no_overlap = False
                break
            seen |= ms

        if no_overlap:
            combined = []
            for _, txs in sources:
                combined.extend(txs)
            combined.sort(key=lambda x: x[0])
            merged[key] = combined
            continue

        # Иначе — есть пересечение. Берём источник с максимальным числом транзакций
        # в пересекающихся месяцах (это многомесячный файл побеждает одномесячный).
        best_source = max(sources, key=lambda s: len(s[1]))
        merged[key] = sorted(best_source[1], key=lambda x: x[0])

    return merged


def merge_sku_to_game(*dicts):
    """Объединяет словари sku -> game (последний выигрывает при конфликтах)."""
    merged = {}
    for d in dicts:
        merged.update(d)
    return merged
