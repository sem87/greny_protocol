import openpyxl
import sys
import os
from logi.logi import logger
# ====================НАЧАЛО НАСТРОЙКИ ====================
# BASE_FILE = "baze/best.xlsx"
# BASE_SHEET_NAME = "title" # None = первый лист, или укажите имя листа
BASE_FILE = "baze/all_data_best.xlsx"
BASE_SHEET_NAME = "mux1" # None = первый лист, или укажите имя листа

# # ====================КОНЕЦ НАСТРОЙКИ ====================


# ==================== НАСТРОЙКА ====================
BASE_FILE = "baze/all_data_best.xlsx"
BASE_SHEET_NAME = "mux1"  # None = первый лист
POINT_COLUMN = 4  # Колонка E (индекс 4, считая с 0) - "Пункт установки"
HEADER_ROW = 3  # Строка с основными заголовками (строка 4 в Excel, индекс 3)
DATA_START_ROW = 6  # Строка начала данных (строка 7 в Excel, индекс 6)


# ================================================


def read_all_data(file_path, sheet_name=None):
    """
    Читает весь лист Excel и возвращает:
    - headers: список заголовков колонок
    - data: список строк (каждая строка - список значений)
    """
    if not os.path.exists(file_path):
        logger.error(f"Файл не найден: {file_path}")
        return None, None

    try:
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)

        if sheet_name is None:
            sheet = wb.active
        else:
            if sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
            else:
                logger.error(f"Лист '{sheet_name}' не найден. Доступны: {wb.sheetnames}")
                wb.close()
                return None, None

        # Читаем заголовки из указанной строки
        headers = []
        for cell in sheet[HEADER_ROW + 1]:  # +1 потому что openpyxl нумерует с 1
            headers.append(str(cell.value) if cell.value else f"Колонка_{cell.column}")

        # Читаем все данные начиная с DATA_START_ROW
        data = []
        for row_idx, row in enumerate(sheet.iter_rows(min_row=DATA_START_ROW + 1, values_only=True),
                                      start=DATA_START_ROW):
            # Пропускаем пустые строки
            if any(cell is not None for cell in row):
                data.append(list(row))

        wb.close()

        logger.info(f"Заголовков: {len(headers)}")
        logger.info(f"Строк данных: {len(data)}")
        logger.info(f"Заголовки: {headers[:10]}...")  # Первые 10 для проверки

        return headers, data

    except Exception as e:
        logger.error(f"Ошибка чтения файла: {e}")
        return None, None


def find_point_by_name(headers, data, point_name, point_column=POINT_COLUMN):
    """
    Ищет строку по названию пункта установки.

    Args:
        headers: список заголовков
        data: список строк
        point_name: название пункта (например, "БАЙМАК")
        point_column: индекс колонки с пунктами установки

    Returns:
        dict: словарь {название_колонки: значение} или None
    """
    if not headers or not data:
        logger.error("Нет данных для поиска")
        return None

    # Нормализуем название для поиска (убираем пробелы, приводим к верхнему регистру)
    point_name_upper = point_name.strip().upper()

    logger.info(f"Ищем пункт: '{point_name_upper}' в колонке {headers[point_column]}")

    for row_idx, row in enumerate(data, start=DATA_START_ROW + 1):
        if point_column < len(row):
            cell_value = row[point_column]
            if cell_value is not None:
                cell_value_str = str(cell_value).strip().upper()

                # Точное совпадение
                if cell_value_str == point_name_upper:
                    logger.info(f"Найден пункт '{point_name}' в строке {row_idx}")

                    # Создаем словарь
                    result = {}
                    for col_idx, header in enumerate(headers):
                        if col_idx < len(row):
                            result[header] = row[col_idx]
                        else:
                            result[header] = None

                    return result

    logger.warning(f"Пункт '{point_name}' не найден")
    return None


def find_all_points(headers, data, point_column=POINT_COLUMN):
    """
    Возвращает список всех пунктов установки.
    """
    points = []
    for row in data:
        if point_column < len(row) and row[point_column] is not None:
            points.append(str(row[point_column]).strip())
    return points


def get_point_value(point_data, column_name):
    """
    Получает значение из данных пункта по имени колонки.
    """
    if not point_data:
        return None

    # Ищем колонку по имени (регистронезависимо)
    column_name_upper = column_name.upper()

    for key, value in point_data.items():
        if key and key.upper() == column_name_upper:
            return value

    # Если не нашли точное совпадение, ищем частичное
    for key, value in point_data.items():
        if key and column_name_upper in key.upper():
            return value

    return None


# ==================== ПРИМЕР ИСПОЛЬЗОВАНИЯ ====================

if __name__ == "__main__":
    logger.info("=" * 60)
    logger.info("Поиск данных по пункту установки")
    logger.info("=" * 60)

    # 1. Читаем файл
    headers, data = read_all_data(BASE_FILE, BASE_SHEET_NAME)

    if not headers or not data:
        logger.error("Не удалось прочитать файл")
        sys.exit(1)

    # 2. Показываем все доступные пункты
    all_points = find_all_points(headers, data)
    logger.info(f"\nВсего пунктов в базе: {len(all_points)}")
    logger.info(f"Первые 10 пунктов: {all_points[:10]}")

    # 3. Ищем конкретный пункт
    search_point = "БАЙМАК"  # Замените на нужный пункт

    logger.info(f"\n--- Поиск пункта: {search_point} ---")
    point_data = find_point_by_name(headers, data, search_point)

    if point_data:
        logger.info(f"\n✅ Данные для пункта '{search_point}':")
        logger.info("-" * 60)

        for header, value in point_data.items():
            if value is not None:  # Показываем только заполненные ячейки
                logger.info(f"  {header}: {value}")

        # Пример получения конкретных значений
        logger.info("\n--- Примеры получения конкретных данных ---")

        # Район
        rayon = get_point_value(point_data, "Район")
        logger.info(f"Район: {rayon}")

        # Широта (градусы)
        shirota_grad = get_point_value(point_data, "Широта WGS 84 (N)")
        logger.info(f"Широта (град): {shirota_grad}")

        # Долгота (градусы)
        dolgota_grad = get_point_value(point_data, "Долгота WGS 84 (E)")
        logger.info(f"Долгота (град): {dolgota_grad}")

        # Мощность
        moshnost = get_point_value(point_data, "Мощность, кВт")
        logger.info(f"Мощность: {moshnost} кВт")

        # Высота подвеса фазовый
        vysota = get_point_value(point_data, "Высота подвеса, фазовый")
        logger.info(f"Высота подвеса: {vysota} м")

    else:
        logger.error(f"Пункт '{search_point}' не найден в базе")
        logger.info("Доступные пункты:")
        for p in all_points[:20]:
            logger.info(f"  - {p}")