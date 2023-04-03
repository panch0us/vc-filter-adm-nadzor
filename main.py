import openpyxl
import datetime

# Проверка перед началом фильтра:
print(f"""                 ВНИМАНИЕ! Перед обработкой таблица должна соответствовать следующим параметрам:

    ******************************************* ОФОРМЛЕНИЕ ТАБЛИЦЫ ************************************************
    Содержать 8 страниц в порядке: [Адреса] [Брянск] [АБДЦ] [АП СООП] [АП ГИБДД] [ЗАПРЕТНИКИ] [ЗАДЕРЖАНИЯ] [ЗАГС]

    ******************************************* ОФОРМЛЕНИЕ СТРАНИЦ ************************************************
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
""")

source_file_name = str(input("Введите название файла для обработки (только формат xlsx): "))

# Если название файл введено без расширения - оно добавляется
if ".xlsx" in source_file_name:
    pass
else:
    source_file_name = source_file_name + ".xlsx"

# Открываем исходный файл *.xlsx и получаем доступ к страницам
try:
    wb_source              = openpyxl.load_workbook(source_file_name)
    sheet_source_bryansk   = wb_source.worksheets[1]
    sheet_source_abdc      = wb_source.worksheets[2]
    sheet_source_soop      = wb_source.worksheets[3]
    sheet_source_gibdd     = wb_source.worksheets[4]
    sheet_source_zapret    = wb_source.worksheets[5]
    sheet_source_zaderj    = wb_source.worksheets[6]
    sheet_source_zags      = wb_source.worksheets[7]
except FileNotFoundError:
    print("Введено неверное название файла или указанный файл отсутсует в текущей дирректории!")
    input("")
    exit()
except IndexError:
    input("Ошибка! В исходном файле не 8 страниц. Проверьте оформление файла!\n"
          "Для выхода нажмите Enter!")
    exit()

# Создаем итоговый файл xlsx для копирования в него строк из исходного файла после фильтра
wb_result              = openpyxl.Workbook()
sheet_result_abdc      = wb_result.create_sheet("АБДЦ", 0)
sheet_result_soop      = wb_result.create_sheet("АП СООП", 1)
sheet_result_gibdd     = wb_result.create_sheet("АП ГИБДД", 2)
sheet_result_zapret    = wb_result.create_sheet("ЗАПРЕТНИКИ", 3)
sheet_result_zaderj    = wb_result.create_sheet("ЗАДЕРЖАНИЯ", 4)
sheet_result_zags      = wb_result.create_sheet("ЗАГС", 5)

# Получаем по 1 строке из кажой страницы для предварительного сравнения строк в страницах
test_bryansk           = list(sheet_source_bryansk.iter_rows(min_row=2, min_col=1, max_col=4, values_only=True))[0]
test_abdc              = list(sheet_source_abdc.iter_rows(   min_row=2, min_col=1, max_col=4, values_only=True))[0]
test_soop              = list(sheet_source_soop.iter_rows(   min_row=1, min_col=1, max_col=4, values_only=True))[0]
test_gibdd             = list(sheet_source_gibdd.iter_rows(  min_row=2, min_col=2, max_col=5, values_only=True))[0]
test_zapret            = list(sheet_source_zapret.iter_rows( min_row=2, min_col=1, max_col=4, values_only=True))[0]
test_zaderj            = list(sheet_source_zaderj.iter_rows( min_row=1, min_col=2, max_col=5, values_only=True))[0]
test_zags              = list(sheet_source_zags.iter_rows(   min_row=2, min_col=2, max_col=5, values_only=True))[0]

try:
    print(f"""
        ***************************** ПРЕДВАРИТЕЛЬНОЕ СРАВНЕНИЕ СТРОК В СТРАНИЦАХ *************************************
        [Страница]\t\t[Фамилия] [Имя] [Отчество] [Дата рождения]
        БРЯНСК:\t\t\t{test_bryansk[0]} {test_bryansk[1]} {test_bryansk[2]} {test_bryansk[3].strftime('%d.%m.%Y')}
        АБДЦ:\t\t\t{test_abdc[0]} {test_abdc[1]} {test_abdc[2]} {str(test_abdc[3])[6:8]}.{str(test_abdc[3])[4:6]}.{str(test_abdc[3])[:4]}
        АП СООП:\t\t{test_soop[0]} {test_soop[1]} {test_soop[2]} {test_soop[3].strftime('%d.%m.%Y')}
        АП ГИБДД:\t\t{test_gibdd[0]} {test_gibdd[1]} {test_gibdd[2]} {test_gibdd[3]}
        ЗАПРЕТНИКИ:\t\t{test_zapret[0]} {test_zapret[1]} {test_zapret[2]} {test_zapret[3].strftime('%d.%m.%Y')}
        ЗАДЕРЖАНИЯ:\t\t{test_zaderj[0]} {test_zaderj[1]} {test_zaderj[2]} {test_zaderj[3].strftime('%d.%m.%Y')}
        ЗАГС:\t\t\t{test_zags[0]} {test_zags[1]} {test_zags[2]} {test_zags[3]}
    """)
except AttributeError:
    input("Исходный файл не соответвсует правилам оформления! Нажмите Enter для выхода.")
    exit()

question = int(input("Если таблица и строки в страницах соответсвтует указанным требованиям - введите цифру 1. "
                     "Для выхода из программы введите любой другой символ: "))
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
            print(f"Совпадение АБДЦ с Брянск: {list_bryansk}")
            sheet_result_abdc.append(row_abdc)

# Фильтр страницы АП СООП
for row_soop in sheet_source_soop.iter_rows(min_row=1, values_only=True):
    list_soop = [cell for cell in row_soop]
    # Приводим формат даты рождения к ДД.ММ.ГГГГ (только если поле не пустое)
    if list_soop[3] != None:
        list_soop[3] = list_soop[3].strftime("%d.%m.%Y")
    #print('ОП СООП:', list_soop)

    # Сравниваем ФИО и дату рождения между АБДЦ и Брянск
    for row_bryansk in sheet_source_bryansk.iter_rows(min_row=2, min_col=1, max_col=4, values_only=True):
        list_bryansk = [cell for cell in row_bryansk]
        # Приводим формат даты рождения к ДД.ММ.ГГГГ (только если поле не пустое)
        if list_bryansk[3] != None:
            list_bryansk[3] = list_bryansk[3].strftime("%d.%m.%Y")
        #print('Bryansk', list_bryansk)
        # Если строка имеет совпадения по ФИО и дате рождения, то сохраяем эту строку в итоговый файл
        if (list_soop[0] == list_bryansk[0] and
                list_soop[1] == list_bryansk[1] and
                list_soop[2] == list_bryansk[2] and
                list_soop[3] == list_bryansk[3]):
            print(f"Совпадение АП СООП с Брянск: {list_bryansk}")
            sheet_result_soop.append(row_soop)

# Фильтр страницы АП ГИБДД
for row_gibdd in sheet_source_gibdd.iter_rows(min_row=2, min_col=2, values_only=True):
    list_gibdd = [cell for cell in row_gibdd]
    # Приводим формат даты рождения к ДД.ММ.ГГГГ (только если поле не пустое)
    #if list_gibdd[3] != None:
    #    list_gibdd[3] = list_gibdd[3]#.strftime("%d.%m.%Y")
    #print('АП ГИБДД:', list_gibdd)

    # Сравниваем ФИО и дату рождения между АБДЦ и Брянск
    for row_bryansk in sheet_source_bryansk.iter_rows(min_row=2, min_col=1, max_col=4, values_only=True):
        list_bryansk = [cell for cell in row_bryansk]
        # Приводим формат даты рождения к ДД.ММ.ГГГГ (только если поле не пустое)
        if list_bryansk[3] != None:
            list_bryansk[3] = list_bryansk[3].strftime("%d.%m.%Y")
        #print('Bryansk', list_bryansk)
        # Если строка имеет совпадения по ФИО и дате рождения, то сохраяем эту строку в итоговый файл
        if (list_gibdd[0] == list_bryansk[0] and
                list_gibdd[1] == list_bryansk[1] and
                list_gibdd[2] == list_bryansk[2] and
                list_gibdd[3] == list_bryansk[3]):
            print(f"Совпадение АП ГИБДД с Брянск: {list_bryansk}")
            sheet_result_gibdd.append(row_gibdd)

# Фильтр страницы ЗАПРЕТНИКИ
for row_zapret in sheet_source_zapret.iter_rows(min_row=2, min_col=1, values_only=True):
    list_zapret = [cell for cell in row_zapret]
    # Приводим формат даты рождения к ДД.ММ.ГГГГ (только если поле не пустое)
    if list_zapret[3] != None:
        list_zapret[3] = list_zapret[3].strftime("%d.%m.%Y")
    #print('ЗАПРЕТНИКИ:', list_zapret)

    # Сравниваем ФИО и дату рождения между АБДЦ и Брянск
    for row_bryansk in sheet_source_bryansk.iter_rows(min_row=2, min_col=1, max_col=4, values_only=True):
        list_bryansk = [cell for cell in row_bryansk]
        # Приводим формат даты рождения к ДД.ММ.ГГГГ (только если поле не пустое)
        if list_bryansk[3] != None:
            list_bryansk[3] = list_bryansk[3].strftime("%d.%m.%Y")
        #print('Bryansk', list_bryansk)
        # Если строка имеет совпадения по ФИО и дате рождения, то сохраяем эту строку в итоговый файл
        if (list_zapret[0] == list_bryansk[0] and
                list_zapret[1] == list_bryansk[1] and
                list_zapret[2] == list_bryansk[2] and
                list_zapret[3] == list_bryansk[3]):
            print(f"Совпадение ЗАПРЕТНИКИ с Брянск: {list_bryansk}")
            sheet_result_zapret.append(row_zapret)

# Фильтр страницы ЗАДЕРЖАНИЯ
for row_zaderj in sheet_source_zaderj.iter_rows(min_row=1, min_col=2, values_only=True):
    list_zaderj = [cell for cell in row_zaderj]
    # Приводим формат даты рождения к ДД.ММ.ГГГГ (только если поле не пустое)
    if list_zaderj[3] != None:
        list_zaderj[3] = list_zaderj[3].strftime("%d.%m.%Y")
    #print('ЗАДЕРЖАНИЯ:', list_zaderj)

    # Сравниваем ФИО и дату рождения между АБДЦ и Брянск
    for row_bryansk in sheet_source_bryansk.iter_rows(min_row=2, min_col=1, max_col=4, values_only=True):
        list_bryansk = [cell for cell in row_bryansk]
        # Приводим формат даты рождения к ДД.ММ.ГГГГ (только если поле не пустое)
        if list_bryansk[3] != None:
            list_bryansk[3] = list_bryansk[3].strftime("%d.%m.%Y")
        #print('Bryansk', list_bryansk)
        # Если строка имеет совпадения по ФИО и дате рождения, то сохраяем эту строку в итоговый файл
        if (list_zaderj[0] == list_bryansk[0] and
                list_zaderj[1] == list_bryansk[1] and
                list_zaderj[2] == list_bryansk[2] and
                list_zaderj[3] == list_bryansk[3]):
            print(f"Совпадение ЗАПРЕТНИКИ с Брянск: {list_bryansk}")
            sheet_result_zaderj.append(row_zaderj)

# Фильтр страницы ЗАГС
for row_zags in sheet_source_zags.iter_rows(min_row=2, min_col=2, values_only=True):
    list_zags = [cell for cell in row_zags]
    # Приводим формат даты рождения к ДД.ММ.ГГГГ (только если поле не пустое)
    #if list_zags[3] != None:
    #    list_zags[3] = list_zags[3]#.strftime("%d.%m.%Y")
    #print('ЗАДЕРЖАНИЯ:', list_zags)

    # Сравниваем ФИО и дату рождения между АБДЦ и Брянск
    for row_bryansk in sheet_source_bryansk.iter_rows(min_row=2, min_col=1, max_col=4, values_only=True):
        list_bryansk = [cell for cell in row_bryansk]
        # Приводим формат даты рождения к ДД.ММ.ГГГГ (только если поле не пустое)
        if list_bryansk[3] != None:
            list_bryansk[3] = list_bryansk[3].strftime("%d.%m.%Y")
        #print('Bryansk', list_bryansk)
        # Если строка имеет совпадения по ФИО и дате рождения, то сохраяем эту строку в итоговый файл
        if (list_zags[0] == list_bryansk[0] and
                list_zags[1] == list_bryansk[1] and
                list_zags[2] == list_bryansk[2] and
                list_zags[3] == list_bryansk[3]):
            print(f"Совпадение ЗАПРЕТНИКИ с Брянск: {list_bryansk}")
            sheet_result_zags.append(row_zags)


result_file_name = "result_" + source_file_name

try:
    wb_result.save(result_file_name)
except PermissionError:
    input("\nОшибка сохранения итогового файла. Возможно предыдущая версия файла уже открыта, закртойте его!\n"
          "Для выхода нажмите Enter!")
    exit()

input("Для завершения программы нажмите Enter ")
