import openpyxl
import sys
import os
from logi.logi import logger
from datetime import datetime, date
from input_data import input_rename
import random

# ====================НАЧАЛО НАСТРОЙКИ ====================
# BASE_FILE = "baze/best.xlsx"
# BASE_SHEET_NAME = "title" # None = первый лист, или укажите имя листа
BASE_FILE = "baze/all_data_best.xlsx"
BASE_SHEET_NAME = "mux1"  # None = первый лист, или укажите имя листа

POINT_COLUMN = 4  # Колонка E (индекс 4, считая с 0) - "Пункт установки"
HEADER_ROW = 3  # Строка с основными заголовками (строка 4 в Excel, индекс 3)
DATA_START_ROW = 6  # Строка начала данных (строка 7 в Excel, индекс 6)


# ====================КОНЕЦ НАСТРОЙКИ ====================
# ============НАЧАЛО ПОЛУЧЕНИЯ ДАННЫХ ИЗ all_data_best.xlsx ======


def format_cell_value(value):
    """Форматирует значение ячейки, конвертируя даты в строки формата DD.MM.YY"""
    # print(f"&&&&&&&&&&&&{value}")
    if value is None:
        return None

    # Если значение - дата или время
    if isinstance(value, (datetime, date)):
        # Форматируем как DD.MM.YY
        return value.strftime('%d.%m.%y')

    # Для всех остальных значений возвращаем как есть
    return value


def read_all_data(file_path, sheet_name=None):
    """Читает весь лист Excel и возвращает:
    - headers: список заголовков колонок
    - data: список строк (каждая строка - список значений)"""
    if not os.path.exists(file_path):
        logger.error(f"Файл не найден: {file_path}")
        return None, None
    try:
        # Убрали data_only=True чтобы корректно читать даты
        wb = openpyxl.load_workbook(file_path, read_only=True)
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
                # Форматируем каждую ячейку в строке (конвертируем даты)
                formatted_row = [format_cell_value(cell) for cell in row]
                data.append(formatted_row)

        wb.close()

        return headers, data

    except Exception as e:
        logger.error(f"Ошибка чтения файла: {e}")
        return None, None


# ====================НАЧАЛО ОТЛАДОЧНЫЙ КОД====================???????????????
# def debug_cell_info(file_path, sheet_name=None, target_row=7):
#     """Отладочная функция для проверки содержимого ячеек"""
#     if not os.path.exists(file_path):
#         logger.error(f"Файл не найден: {file_path}")
#         return
#
#     wb = openpyxl.load_workbook(file_path, read_only=True)
#     if sheet_name is None:
#         sheet = wb.active
#     else:
#         sheet = wb[sheet_name]
#
#     print(f"\n{'=' * 80}")
#     print(f"ОТЛАДКА: Строка {target_row} (индекс {target_row - 1})")
#     print(f"{'=' * 80}")
#
#     # Получаем строку
#     row = list(sheet.iter_rows(min_row=target_row, max_row=target_row, values_only=False))[0]
#
#     for cell in row:
#         value = cell.value
#         col_idx = cell.column - 1  # 0-based индекс
#
#         # Показываем только непустые ячейки или ячейки с датами
#         if value is not None or 'дата' in str(cell.value).lower():
#             print(f"Колонка {col_idx} ({cell.coordinate}):")
#             print(f"  Значение: {repr(value)}")
#             print(f"  Тип: {type(value).__name__}")
#             print(f"  Числовой формат: {cell.number_format}")
#             print()
#
#     wb.close()
#
#
# def read_all_data_with_debug(file_path, sheet_name=None, debug_row=7):
#     """Читает файл с отладкой"""
#     # Сначала запускаем отладку
#     debug_cell_info(file_path, sheet_name, debug_row)
#
#     # Затем читаем данные как обычно
#     return read_all_data(file_path, sheet_name)

# ====================КОНЕЦ ОТЛАДОЧНЫЙ КОД====================????????????????


def find_point_by_name(headers, data, point_name, point_column=POINT_COLUMN):
    """Ищет строку по названию пункта установки.
    Args:
        headers: список заголовков
        data: список строк
        point_name: название пункта (например, "БАЙМАК")
        point_column: индекс колонки с пунктами установки
    Returns:
        dict: словарь {название_колонки: значение} или None"""
    if not headers or not data:
        logger.error("Нет данных для поиска")
        return None

    # Нормализуем название для поиска (убираем пробелы, приводим к верхнему регистру)
    point_name_upper = point_name.strip().upper()
    # logger.info(f"Ищем пункт: '{point_name_upper}' в колонке {headers[point_column]}")
    for row_idx, row in enumerate(data, start=DATA_START_ROW + 1):
        if point_column < len(row):
            cell_value = row[point_column]
            if cell_value is not None:
                cell_value_str = str(cell_value).strip().upper()
                # Точное совпадение
                if cell_value_str == point_name_upper:
                    # logger.info(f"Найден пункт '{point_name}' в строке {row_idx}")
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


def get_point_value(point_data, column_name):
    """Получает значение из данных пункта по имени колонки."""
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


# ============КОНЕЦ ПОЛУЧЕНИЯ ДАННЫХ ИЗ all_data_best.xlsx ======
if __name__ == "__main__":
    # # =========ОТЛАДКА
    # headers, data = read_all_data_with_debug(BASE_FILE, BASE_SHEET_NAME, debug_row=26)
    # # =========ОТЛАДКА

    # =======Начало входные данные=====
    itog_number = None
    search_point = "БАЙки"  # Место где стоит вышка
    location_measure_metrics = "УкАрлино" # Граничный населенный пункт где беруться замеры
    date_protocol = "10.06.2016"
    # =======Конец входные данные======

    # 1. Читаем файл
    headers, data = read_all_data(BASE_FILE, BASE_SHEET_NAME)
    # 3. Ищем конкретный пункт
    logger.info(f"Поиск пункта: {search_point}")
    point_data = find_point_by_name(headers, data, search_point)
    print(point_data)

    # ПАРАМЕТРЫ
    # номер протокола
    if itog_number == None:
        itog_number = random.randint(100, 999)

    # район
    rayon = get_point_value(point_data, "Район")
    itog_rayon = search_point.strip().capitalize() + f"({rayon} р-н)"
    logger.info(f"Район: {itog_rayon}")

    # номер разрешения
    razrecshenie = get_point_value(point_data, "Разрешение на использование радиочастот")
    razrecshenie_data_vidachi = get_point_value(point_data, "Дата выдачи")
    itog_razrecshenie = f"Разрешение " + razrecshenie + " от " + str(razrecshenie_data_vidachi)
    logger.info(f"Разрешение: {itog_razrecshenie}")

    # координаты
    coordinates_shirota_1 = get_point_value(point_data, "Широта WGS 84 (N)")
    coordinates_shirota_2 = get_point_value(point_data, "Колонка_9")
    coordinates_shirota_3 = get_point_value(point_data, "Колонка_10")
    coordinates_dolgota_1 = get_point_value(point_data, "Долгота WGS 84 (E)")
    coordinates_dolgota_2 = get_point_value(point_data, "Колонка_12'")
    coordinates_dolgota_3 = get_point_value(point_data, "Колонка_13")
    # Правильный формат: 55° 45' 20.88" N, 37° 37' 2.28" E
    itog_coordinates = f"{coordinates_shirota_1}° {coordinates_shirota_2}' {coordinates_shirota_3}\" N, {coordinates_dolgota_1}° {coordinates_dolgota_2}' {coordinates_dolgota_3}\" E"

    # мощность передатчика
    power = get_point_value(point_data, "Мощность, кВт")
    itog_power = f"{power}, кВт"

    # высота подвеса
    height = get_point_value(point_data, "Высота подвеса над уровнем земли (из ЧТП)")
    itog_height = f"{height}, м"

    # средство измерений
    svidetelstvo_izmerenia = {"16.02.2027": ["17.02.2026", "С-БЕВ/17-02-2026/506326766"],
                              "20.01.2026": ["21.01.2025", "С-БЕВ/21-01-2025/404395610"],
                              "27.12.2024": ["28.12.2023", "С-БЕВ/28-12-2023/306877891"],
                              "08.09.2023": ["09.09.2022", "С-БЕВ/09-09-2022/187120735"],
                              "04.08.2022": ["05.08.2021", "С-БЕВ/05-08-2021/85940153"],
                              "12.07.2021": ["13.07.2020", "УС-20/430"],
                              "03.07.2020": ["04.07.2019", "УС-1864/19"],
                              "13.06.2019": ["14.06.2018", "8360/2018-717068"],
                              "21.06.2018": ["22.06.2017", "5057/2017-713547"],
                              "19.06.2017": ["20.06.2016", "2854/2016-710188"]}

    protocol_date = datetime.strptime(date_protocol, "%d.%m.%Y")
    # Ищем минимальный ключ, который больше date_protocol
    result_key = min((key for key in svidetelstvo_izmerenia if datetime.strptime(key, "%d.%m.%Y") > protocol_date),
                     key=lambda k: datetime.strptime(k, "%d.%m.%Y"))
    svidetelstvo_poverka_date = svidetelstvo_izmerenia[result_key][0]
    svidetelstvo_poverka_number = svidetelstvo_izmerenia[result_key][1]




    # Берем граничный населенный пункт и парсим координаты
    itog_location_measure_metrics = location_measure_metrics.strip().capitalize()





    # запись
    input_rename(itog_rayon=itog_rayon, itog_razrecshenie=itog_razrecshenie, date_protocol=date_protocol,
                 itog_number=itog_number, itog_coordinates=itog_coordinates, itog_power=itog_power,
                 itog_height=itog_height, svidetelstvo_poverka_date=svidetelstvo_poverka_date,
                 svidetelstvo_poverka_number=svidetelstvo_poverka_number,itog_location_measure_metrics=itog_location_measure_metrics)
