import openpyxl
import datetime

# Открываем исходный файл *.xlsx
wb_source              = openpyxl.load_workbook('20230315.xlsx')
sheet_source_bryansk   = wb_source.worksheets[1]
sheet_source_abdc      = wb_source.worksheets[2]
sheet_source_soop      = wb_source.worksheets[3]
sheet_source_gibdd     = wb_source.worksheets[4]
sheet_source_zapret    = wb_source.worksheets[5]
sheet_source_zaderj    = wb_source.worksheets[6]
sheet_source_zags      = wb_source.worksheets[7]

# Создаем итоговый файл xlsx для копирования в него строк из исходного файла после фильтра
wb_result              = openpyxl.Workbook()
sheet_result_abdc      = wb_result.create_sheet("АБДЦ", 0)
sheet_result_soop      = wb_result.create_sheet("АП СООП", 1)
sheet_result_gibdd     = wb_result.create_sheet("АП ГИБДД", 2)
sheet_result_zapret    = wb_result.create_sheet("ЗАПРЕТНИКИ", 3)
sheet_result_zaderj    = wb_result.create_sheet("ЗАДЕРЖАНИЯ", 4)
sheet_result_zags      = wb_result.create_sheet("ЗАГС", 5)

# Проверка перед началом фильтра:
print(f"""           ВНИМАНИЕ! Перед обработкой таблица должна соответствовать следующим параметрам:
    ---------------------------------------------------------------------------------------------------------------
    БРЯНСК:     Заголовок:      [Фамилия] [Имя] [Отчество] [Дата рождения (формат: ДД.ММ.ГГГГ)]
    ---------------------------------------------------------------------------------------------------------------
    АБДЦ:       Заголовок:      [Фамилия] [Имя] [Отчество] [Дата рождения (формат: ГГГГММДД)]
    ---------------------------------------------------------------------------------------------------------------
    АП СООП:    1 Строка:       [Фамилия] [Имя] [Отчество] [Дата рождения (формат: ДД.ММ.ГГГГ)] (!!! ЗАГОЛОВКА НЕТ)
    ---------------------------------------------------------------------------------------------------------------
    АП ГИБДД:   Заголовок:  [№] [Фамилия] [Имя] [Отчество] [Дата рождения (формат: ДД.ММ.ГГГГ)]
    ---------------------------------------------------------------------------------------------------------------
    ЗАПРЕТНИКИ: Заголовок:      [Фамилия] [Имя] [Отчество] [Дата рождения (формат: ДД.ММ.ГГГГ)]
    ---------------------------------------------------------------------------------------------------------------
    ЗАДЕРЖАНИЯ: 1 Строка:   [№] [Фамилия] [Имя] [Отчество] [Дата рождения (формат: ДД.ММ.ГГГГ)] (!!! ЗАГОЛОВКА НЕТ)
    ---------------------------------------------------------------------------------------------------------------
    ЗАГС:       Заголовок:  [№] [Фамилия] [Имя] [Отчество] [Дата рождения (формат: ДД.ММ.ГГГГ)]
    ---------------------------------------------------------------------------------------------------------------
""")

question = int(input("Введите 1, если таблица соответсвтует указанным требованиям. Если нет, введите 2: "))

if question != 1:
    exit()

# Фильтр страницы АБДЦ
for row_abdc in sheet_source_abdc.iter_rows(min_row=2, values_only=True):
    list_abdc = [cell for cell in row_abdc]
    # Приводим формат даты рождения к ДД.ММ.ГГГГ (только если поле не пустое)
    if list_abdc[3] != None:
        list_abdc[3] = str(list_abdc[3])[6:8] + '.' + str(list_abdc[3])[4:6] + '.' + str(list_abdc[3])[:4]
    #print('АДБЦ:', list_abdc)

    # Сравниваем ФИО и дату рождения между АБДЦ и Брянск
    for row_bryansk in sheet_source_bryansk.iter_rows(min_row=2, min_col=1, max_col=4, values_only=True):
        list_bryansk = [cell for cell in row_bryansk]
        # Приводим формат даты рождения к ДД.ММ.ГГГГ (только если поле не пустое)
        if list_bryansk[3] != None:
            list_bryansk[3] = list_bryansk[3].strftime("%d.%m.%Y")
        #print('Bryansk', list_bryansk)
        # Если строка имеет совпадения по ФИО и дате рождения, то сохраяем эту строку в итоговый файл
        if (list_abdc[0] == list_bryansk[0] and
                list_abdc[1] == list_bryansk[1] and
                list_abdc[2] == list_bryansk[2] and
                list_abdc[3] == list_bryansk[3]):
            print(f"Совпадение АБДЦ с Брянск: {list_bryansk[0], list_bryansk[1], list_bryansk[2], list_bryansk[3]}")
            sheet_result_abdc.append(row_abdc)


wb_result.save('result.xlsx')