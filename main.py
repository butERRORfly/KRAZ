import openpyxl

# Получаем данные из файла
wb = openpyxl.load_workbook('DAM_KEP.xlsx')

# Получаем доступ к активному листу
ws = wb.active

date_question = input

# Фильтрация по ключам
# Создаем новые листы
good_sheet = wb.create_sheet('Good')
bad_sheet = wb.create_sheet('Bad')
none_sheet = wb.create_sheet('None')
date_sheet = wb.create_sheet('DateChoose')

# Копируем названия столбцов на каждый новый лист
main_titles_values_good = [good_sheet.append(row) for row in ws.iter_rows(min_row=1, max_row=1, values_only=True)]
main_titles_values_bad = [bad_sheet.append(row) for row in ws.iter_rows(min_row=1, max_row=1, values_only=True)]
main_titles_values_none = [none_sheet.append(row) for row in ws.iter_rows(min_row=1, max_row=1, values_only=True)]
main_titles_values_date = [date_sheet.append(row) for row in ws.iter_rows(min_row=1, max_row=1, values_only=True)]

# Диапазон ячеек, на котором мы ищем флаги "Good", "Bad" и "None"
search_value_range = ws['D2':'D2119']
for i in search_value_range:

    # Обрабатываем флаг "Good"
    if i[0].value == 'Good':
        # Добавляем значения ячеек на новый лист "Good"
        good_values = [good_sheet.append(row) for row in ws.iter_rows(min_row=i[0].row, max_row=i[-1].row, values_only=True)]

    # Обрабатываем флаг "Bad"
    if i[0].value == 'Bad':
        bad_values = [bad_sheet.append(row) for row in ws.iter_rows(min_row=i[0].row, max_row=i[-1].row, values_only=True)]

# Диапазон ячеек, на котром мы ищем требуемую дату поиска
search_date_range = ws['B2':'B2119']
for el in search_date_range:
    filter_date_search = 'Aug 18'
    if filter_date_search in el[0].value:
        date_values = [date_sheet.append(row) for row in ws.iter_rows(min_row=el[0].row, max_row=el[-1].row, values_only=True)]


# Сохраняем наши изменения
wb.save('DAM_KEP.xlsx')
