def input_rename(itog_rayon, itog_razrecshenie):
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
    sheet['G5'] = date_protocol

    # 4. Сохраняем файл по новому пути
    workbook.save(output_path)
    print(f"✅ Файл успешно сохранен по пути:\n{output_path}")