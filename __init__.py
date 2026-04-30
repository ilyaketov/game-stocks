"""
Оркестратор для облачной версии — работает с in-memory данными.

Главные функции:
  process_uploaded_files(files): принимает список UploadedFile из st.file_uploader,
    парсит их, объединяет транзакции, возвращает структуру с месяцами.
  build_report_for_month(year, month, data): возвращает BytesIO с готовым xlsx.
"""
from datetime import datetime
from collections import defaultdict
from io import BytesIO

from .parser import parse_inventory_report, get_months_in_data
from .calculator import build_flat, month_period, merge_transactions, merge_sku_to_game
from .report_writer import write_report


def process_uploaded_files(uploaded_files):
    """
    Обрабатывает загруженные пользователем файлы.

    uploaded_files: список объектов с атрибутами .name, .getvalue() (как у Streamlit
                    UploadedFile) либо просто dicts {'name', 'bytes'}.

    Возвращает dict:
      {
        'transactions': объединённый словарь транзакций,
        'sku_to_game': объединённый словарь sku → название игры,
        'all_periods': отсортированный список (year, month) представленных в данных,
        'source_bytes_by_period': dict[(y,m)] -> bytes — какие исходные байты
                                  использовать для листа «Источник» в отчёте за этот месяц,
        'file_summary': список {'filename', 'periods'} — что прочитали из каждого файла,
      }
    """
    if not uploaded_files:
        return None

    # Нормализуем входные объекты к (name, bytes)
    normalized = []
    for f in uploaded_files:
        if hasattr(f, 'getvalue') and hasattr(f, 'name'):
            normalized.append((f.name, f.getvalue()))
        elif isinstance(f, dict):
            normalized.append((f['name'], f['bytes']))
        else:
            raise ValueError("Неподдерживаемый формат файла в uploaded_files")

    all_txs = []
    all_sku_maps = []
    file_summary = []
    # Список (priority, name, bytes, periods_in_file). Меньше priority = выше:
    #   0 — одномесячный (приоритетный для своего месяца)
    #   1 — многомесячный (используется только если нет одномесячного)
    files_data = []

    for name, raw in normalized:
        bio = BytesIO(raw)
        try:
            txs, sku_map = parse_inventory_report(bio)
        except Exception as e:
            raise ValueError(f"Не удалось прочитать «{name}»: {e}") from e

        periods = get_months_in_data(txs)
        priority = 0 if len(periods) == 1 else 1

        all_txs.append(txs)
        all_sku_maps.append(sku_map)
        files_data.append((priority, name, raw, periods))
        file_summary.append({'filename': name, 'periods': periods})

    if not all_txs:
        return None

    transactions = merge_transactions(*all_txs)
    sku_to_game = merge_sku_to_game(*all_sku_maps)

    # Какой исходник использовать для каждого месяца:
    # сначала проходим одномесячные (priority=0), потом многомесячные (priority=1).
    source_bytes_by_period = {}
    for priority, name, raw, periods in sorted(files_data, key=lambda x: x[0]):
        for p in periods:
            if p not in source_bytes_by_period:
                source_bytes_by_period[p] = raw

    all_periods = sorted({p for periods in [d[3] for d in files_data] for p in periods})

    return {
        'transactions': transactions,
        'sku_to_game': sku_to_game,
        'all_periods': all_periods,
        'source_bytes_by_period': source_bytes_by_period,
        'file_summary': file_summary,
    }


def build_report_for_month(year, month, data):
    """
    Строит отчёт за указанный месяц на основе data (результат process_uploaded_files).

    Возвращает BytesIO с готовым xlsx.
    """
    period_start, period_end = month_period(year, month)
    flat = build_flat(
        data['transactions'],
        data['sku_to_game'],
        period_start,
        period_end,
    )
    source_bytes = data['source_bytes_by_period'].get((year, month))
    return write_report(period_start, period_end, flat, source_bytes, output_path=None), flat


def report_filename_for_period(year, month):
    from calendar import monthrange
    last_day = monthrange(year, month)[1]
    return f"Стоки игр на {last_day:02d}.{month:02d}.{year}.xlsx"
