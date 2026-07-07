import openpyxl
import os

# ================= НАСТРОЙКИ =================
# Путь к исходному файлу (должен существовать)
input_file = 'baze/best.xlsx'

# Папка, в которую нужно сохранить результат
output_folder = 'posle'


# =============================================

def input_rename(itog_rayon, itog_razrecshenie, date_protocol, itog_number, itog_coordinates, itog_power, itog_height,
                 svidetelstvo_poverka_number, svidetelstvo_poverka_date, itog_location_measure_metrics,
                 point_geo_1_shirota, point_geo_1_dolgota, point_geo_2_shirota, point_geo_2_dolgota,
                 point_geo_3_shirota, point_geo_3_dolgota, point_geo_4_shirota, point_geo_4_dolgota,
                 point_geo_5_shirota, point_geo_5_dolgota, point_geo_6_shirota, point_geo_6_dolgota):
    # Формируем полный путь для сохранения
    # Имя нового файла
    output_filename = f"отчет_{itog_rayon}_{itog_razrecshenie}.xlsx"
    output_path = os.path.join(output_folder, output_filename)

    # 1. Создаем папку для сохранения, если она еще не существует
    os.makedirs(output_folder, exist_ok=True)

    # 2. Открываем существующий Excel файл
    try:
        workbook = openpyxl.load_workbook(input_file)
    except FileNotFoundError:
        print(f"Ошибка: Файл '{input_file}' не найден!")
        exit()

    # Выбираем лист для редактирования
    # (можно указать имя листа: workbook['Лист1'] или оставить workbook.active для текущего)
    sheet = workbook.active

    # 3. Вставляем данные
    # --- Вариант А: Вставка в конкретные ячейки по координатам ---
    sheet['N7'] = itog_number
    sheet['G5'] = date_protocol
    sheet['R7'] = date_protocol
    sheet['N12'] = itog_rayon
    sheet['E17'] = itog_razrecshenie
    sheet['Q21'] = itog_coordinates
    sheet['Q22'] = itog_power
    sheet['Q23'] = itog_height
    sheet['M35'] = svidetelstvo_poverka_number
    sheet['M36'] = svidetelstvo_poverka_date
    sheet['E53'] = itog_location_measure_metrics
    #=======Широта долгота=======
    sheet['H53'] = point_geo_1_shirota
    sheet['I53'] = point_geo_1_dolgota
    sheet['H54'] = point_geo_2_shirota
    sheet['I54'] = point_geo_2_dolgota
    sheet['H55'] = point_geo_3_shirota
    sheet['I55'] = point_geo_3_dolgota
    sheet['H56'] = point_geo_4_shirota
    sheet['I56'] = point_geo_4_dolgota
    sheet['H57'] = point_geo_5_shirota
    sheet['I57'] = point_geo_5_dolgota
    sheet['H58'] = point_geo_6_shirota
    sheet['I58'] = point_geo_6_dolgota
    # # --- Вариант Б: Вставка массива данных (например, в 4-ю строку) ---
    # new_data = ['Иванов', 'Иван', 'Иванович', 35, 'Менеджер']
    # start_row = 4
    # for col_index, value in enumerate(new_data, start=1): # start=1 означает колонку A
    #     sheet.cell(row=start_row, column=col_index, value=value)

    # 4. Сохраняем файл по новому пути
    workbook.save(output_path)
    print(f"✅ Файл успешно сохранен по пути:\n{output_path}")


if __name__ == "__main__":
    pass
