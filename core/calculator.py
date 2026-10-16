"""Расчёт плоской таблицы остатков за указанный календарный период."""
from datetime import datetime
from calendar import monthrange
from collections import defaultdict


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
    # Магазины, у которых есть хотя бы одна закупка (Bill) за всю историю до period_end.
    # Магазин БЕЗ закупок — это склад-накопитель (сток приходит только перемещениями),
    # он исключается из анализа рисков (см. report_writer).
    stores_with_bills = set()
    for (sku, store), txs in transactions.items():
        for t in txs:
            dt, ttype = t[0], t[1]
            if dt <= period_end and ttype == 'Bill':
                stores_with_bills.add(store)
                if sku not in last_bill_by_sku or dt > last_bill_by_sku[sku]:
                    last_bill_by_sku[sku] = dt

    # --- Классификация прочих движений на ПЕРЕВОДЫ и КОРРЕКТИРОВКИ по парам Num ---
    # Перевод между магазинами оформляется в QuickBooks двумя встречными документами
    # (напр. Num 3304: −N из магазина А, Num 3305: +N в магазин Б). Документ опознаётся
    # как перевод, если у него есть парный документ с ТОЧНО противоположным профилем
    # (по каждому артикулу qty с обратным знаком). Остальные документы (одиночные
    # списания/оприходования, Vendor Credit и т.п.) — корректировки, меняющие общий
    # остаток. Так переводы всегда взаимозачитываются в ноль, а корректировки — нет.
    def _num(t):
        return t[6] if len(t) > 6 else ''
    _period_docs = defaultdict(lambda: defaultdict(float))  # num -> {sku: qty}
    for (sku, store), txs in transactions.items():
        for t in txs:
            if period_start <= t[0] <= period_end and t[1] not in ('Bill', 'Invoice') and _num(t):
                _period_docs[_num(t)][sku] += t[2]

    def _profile(num):
        return tuple(sorted((s, round(q, 3)) for s, q in _period_docs[num].items() if abs(q) > 1e-6))

    _by_prof = defaultdict(list)
    for num in _period_docs:
        p = _profile(num)
        if p:
            _by_prof[p].append(num)

    transfer_nums = set()
    _used = set()
    for num in _period_docs:
        if num in _used:
            continue
        p = _profile(num)
        if not p:
            continue
        want = tuple(sorted((s, -q) for s, q in p))
        cands = [m for m in _by_prof.get(want, []) if m not in _used and m != num]
        if cands:
            m = cands[0]
            _used.add(num)
            _used.add(m)
            transfer_nums.add(num)
            transfer_nums.add(m)

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
        # Прочие движения (перемещения/корректировки) = всё, что НЕ Bill и НЕ Invoice:
        # Inventory Adjust, Inventory Qty Adjust, Vendor Credit и любые будущие типы.
        # Раньше привязывались к точной строке 'Inventory Qty Adjust' — но QuickBooks
        # в разных выгрузках называет тип по-разному ('Inventory Adjust' без «Qty»),
        # из-за чего корректировки терялись в «Перемещениях», оставаясь в конечном
        # остатке, и начальный остаток уходил в минус. Классификация «всё прочее»
        # делает тождество закуп − продано + перемещения = конец всегда верным и
        # устойчивым к переименованию типов.
        _is_move = lambda t: t[1] not in ('Bill', 'Invoice')
        # Переводы — движения, чей документ Num опознан как парный (см. transfer_nums).
        _is_transfer = lambda t: _is_move(t) and _num(t) in transfer_nums
        # Корректировки — все прочие движения (одиночные списания/оприходования,
        # Vendor Credit, движения без Num и т.п.).
        _is_corr = lambda t: _is_move(t) and _num(t) not in transfer_nums

        transfer_qty = sum(t[2] for t in txs_in_period if _is_transfer(t))
        transfer_val = sum(t[3] for t in txs_in_period if _is_transfer(t))
        corr_qty = sum(t[2] for t in txs_in_period if _is_corr(t))
        corr_val = sum(t[3] for t in txs_in_period if _is_corr(t))

        start_qty = end_qty - bills_qty + invoices_qty - transfer_qty - corr_qty
        start_val = end_val - bills_val + invoices_val - transfer_val - corr_val

        days_since = None
        if sku in last_bill_by_sku:
            days_since = (period_end - last_bill_by_sku[sku]).days

        has_movement = len(txs_in_period) > 0
        has_any_sale_ever = any(t[1] == 'Invoice' for t in txs_until_end)

        flat.append({
            'store': store,
            'game': sku_to_game.get(sku, ''),
            'sku': sku,
            'start_qty': start_qty,
            'sold_qty': invoices_qty,
            'bought_qty': bills_qty,
            'transfer_qty': transfer_qty,
            'corr_qty': corr_qty,
            'end_qty': end_qty,
            'start_val': round(start_val, 2),
            'sold_val': round(invoices_val, 2),
            'bought_val': round(bills_val, 2),
            'transfer_val': round(transfer_val, 2),
            'corr_val': round(corr_val, 2),
            'end_val': round(end_val, 2),
            'days_since_bill': days_since,
            'has_movement_in_period': has_movement,
            'has_any_sale_ever': has_any_sale_ever,
            'is_warehouse': store not in stores_with_bills,
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
