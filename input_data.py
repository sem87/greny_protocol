import openpyxl
import os
import random
import shutil
import subprocess
import tempfile
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
                 point_geo_5_shirota, point_geo_5_dolgota, point_geo_6_shirota, point_geo_6_dolgota, azimut_raschetni_1,
                 azimut_raschetni_2, azimut_raschetni_3, azimut_raschetni_4, azimut_raschetni_5, azimut_raschetni_6,
                 azimut_izmereni_1, azimut_izmereni_2, azimut_izmereni_3, azimut_izmereni_4, azimut_izmereni_5,
                 azimut_izmereni_6, gauss_value, cell_id, itog_koeff_ysilenia, itog_type_anten,
                 itog_prozent_ohvata_naselenia):
    # Формируем полный путь для сохранения
    # Имя нового файла
    output_filename = f"Протокол_зоны_НЦТВ_{itog_rayon}_{itog_location_measure_metrics}_{date_protocol}.xlsx"
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
    # sheet['G5'] = date_protocol
    sheet['P7'] = date_protocol
    sheet['N12'] = itog_rayon
    sheet['E17'] = itog_razrecshenie
    sheet['M21'] = itog_coordinates
    sheet['M22'] = itog_power
    sheet['M23'] = itog_height
    sheet['M24'] = itog_type_anten
    sheet['M25'] = itog_koeff_ysilenia
    sheet['M35'] = svidetelstvo_poverka_number
    sheet['M36'] = svidetelstvo_poverka_date
    sheet['E53'] = itog_location_measure_metrics
    # =======Широта долгота=======
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
    sheet['L53'] = azimut_raschetni_1
    sheet['L54'] = azimut_raschetni_2
    sheet['L55'] = azimut_raschetni_3
    sheet['L56'] = azimut_raschetni_4
    sheet['L57'] = azimut_raschetni_5
    sheet['L58'] = azimut_raschetni_6
    sheet['M53'] = azimut_izmereni_1
    sheet['M54'] = azimut_izmereni_2
    sheet['M55'] = azimut_izmereni_3
    sheet['M56'] = azimut_izmereni_4
    sheet['M57'] = azimut_izmereni_5
    sheet['M58'] = azimut_izmereni_6
    sheet['O53'] = gauss_value
    sheet['N53'] = gauss_value + random.randint(20, 40) / 10.0
    sheet['O54'] = gauss_value
    sheet['N54'] = gauss_value + random.randint(20, 45) / 10.0
    sheet['O55'] = gauss_value
    sheet['N55'] = gauss_value + random.randint(30, 40) / 10.0
    sheet['O56'] = gauss_value
    sheet['N56'] = gauss_value + random.randint(20, 30) / 10.0
    sheet['O57'] = gauss_value
    sheet['N57'] = gauss_value + random.randint(20, 45) / 10.0
    sheet['O58'] = gauss_value
    sheet['N58'] = gauss_value + random.randint(20, 40) / 10.0
    sheet['K53'] = cell_id
    sheet['K54'] = cell_id
    sheet['K55'] = cell_id
    sheet['K56'] = cell_id
    sheet['K57'] = cell_id
    sheet['K58'] = cell_id
    sheet['T53'] = itog_prozent_ohvata_naselenia
    sheet['T54'] = itog_prozent_ohvata_naselenia
    sheet['T55'] = itog_prozent_ohvata_naselenia
    sheet['T56'] = itog_prozent_ohvata_naselenia
    sheet['T57'] = itog_prozent_ohvata_naselenia
    sheet['T58'] = itog_prozent_ohvata_naselenia
    # sheet['J78'] = output_filename

    # 4. Сохраняем файл по новому пути
    workbook.save(output_path)
    print(f"✅ Файл успешно сохранен по пути:\n{output_path}")


def convert_xlsx_to_pdf(folder_name):
    """Конвертирует все файлы .xlsx в .pdf в указанной папке
    с помощью LibreOffice в headless режиме."""
    # Проверяем, существует ли папка
    if not os.path.isdir(folder_name):
        print(f"Ошибка: Папка '{folder_name}' не найдена.")
        return
    # Получаем абсолютный путь (LibreOffice лучше работает с абсолютными путями)
    abs_folder_path = os.path.abspath(folder_name)
    # # Счетчик для статистики
    # converted_count = 0

    # Проходим по всем файлам в папке
    for filename in os.listdir(abs_folder_path):
        # Проверяем, что файл имеет расширение .xlsx
        if filename.lower().endswith('.xlsx'):
            input_file_path = os.path.join(abs_folder_path, filename)
            # Формируем команду для вызова LibreOffice
            command = [
                'libreoffice',
                '--headless',  # Запуск без графического интерфейса
                '--norestore',  # Не восстанавливать предыдущую сессию
                '--convert-to', 'pdf',  # Формат конвертации
                '--outdir', abs_folder_path,  # Папка для сохранения (та же самая)
                input_file_path  # Путь к исходному файлу
            ]
            #
            # print(f"Конвертирую: {filename} ...", end=" ")
            try:
                # Запускаем процесс.
                # stdout и stderr перенаправляем в DEVNULL, чтобы не засорять консоль
                subprocess.run(
                    command,
                    check=True,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL
                )
                print(f"Готово! (Сохранено как {filename.replace('.xlsx', '.pdf')})")
            except subprocess.CalledProcessError:
                print("Ошибка при конвертации.")
            except FileNotFoundError:
                print("\nКритическая ошибка: LibreOffice не найден в системе!")
                print("Установите его: sudo apt install libreoffice-calc")
                break


if __name__ == "__main__":
    convert_xlsx_to_pdf(folder_name=output_folder)
