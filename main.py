import openpyxl

# Получаем данные из файла
wb = openpyxl.load_workbook('DAM_KEP.xlsx')

# Получаем доступ к активному листу
ws = wb.active

# Фильтрация по ключам
# Создаем новые листы
good_sheet = wb.create_sheet('Good')
bad_sheet = wb.create_sheet('Bad')
none_sheet = wb.create_sheet('None')


# Копируем названия столбцов на каждый новый лист
main_titles_values_good = [good_sheet.append(row) for row in ws.iter_rows(min_row=1, max_row=1, values_only=True)]
main_titles_values_bad = [bad_sheet.append(row) for row in ws.iter_rows(min_row=1, max_row=1, values_only=True)]
main_titles_values_none = [none_sheet.append(row) for row in ws.iter_rows(min_row=1, max_row=1, values_only=True)]

# Диапазон ячеек, на котором мы ищем флаги "Good", "Bad" и "None"
search_range = ws['D2':'D2119']
for i in search_range:

    # Обрабатываем флаг "Good"
    if i[0].value == 'Good':
        # Добавляем значения ячеек на новый лист "Good"
        good_values = [good_sheet.append(row) for row in ws.iter_rows(min_row=i[0].row, max_row=i[-1].row, values_only=True)]

    # Обрабатываем флаг "Bad"
    if i[0].value == 'Bad':
        bad_values = [bad_sheet.append(row) for row in ws.iter_rows(min_row=i[0].row, max_row=i[-1].row, values_only=True)]

# Сохраняем наши изменения
wb.save('DAM_KEP.xlsx')
