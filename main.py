import openpyxl

nums_sheets = [2, 3, 4, 5, 6, 7]


def open_sheets(nums_sheets: list):
    for num_sheet in nums_sheets:
        sheet = wb.worksheets[num_sheet]
        yield sheet


def filter_sheets(sheets: object):
    sheet_bryansk = wb.worksheets[1]
    print('Количество строк в Брянск:', sheet_bryansk.max_row)

    for sheet in sheets:
        print(f'Начало обработки страницы: {sheet}')

        if "АБДЦ" in str(sheet):
            print('Количество строк в АБДЦ:', sheet.max_row)
            count_1 = 2 # Подсчет для первого цикла for
            count_2 = 2 # Подсчет для второго цикла for

            for row in range(sheet.max_row - 1):
                A1 = 'A' + str(count_1)
                B1 = 'B' + str(count_1)
                C1 = 'C' + str(count_1)
                D1 = 'D' + str(count_1)
                #print('АБДЦ: ', A1)

                surname_1  = str(sheet[A1].value).strip().upper()
                name_1     = str(sheet[B1].value).strip().upper()
                otch_1     = str(sheet[C1].value).strip().upper()
                birthday_1 = str(sheet[D1].value).strip().upper()

                for row in range(sheet_bryansk.max_row - 1):
                    A2 = 'A' + str(count_2)
                    B2 = 'B' + str(count_2)
                    C2 = 'C' + str(count_2)
                    D2 = 'D' + str(count_2)

                    surname_2  = str(sheet_bryansk[A2].value).strip().upper()
                    name_2     = str(sheet_bryansk[B2].value).strip().upper()
                    otch_2     = str(sheet_bryansk[C2].value).strip().upper()
                    birthday_2 = str(sheet_bryansk[D2].value).strip().upper()

                    #print('Брянск: ', A1)

                    if (surname_1 == surname_2) and (name_1 == name_2) and (otch_1 == otch_2):
                        print(f'Совпадение: {count_1}', surname_1, name_1, otch_1, birthday_1)
                    else:
                        for cell in sheet[count_1]:
                            cell.value = None

                    count_2 = count_2 + 1

                count_2 = 2
                count_1 = count_1 + 1


        elif "АП СООП" in str(sheet):
            print("АП СООП")


if __name__ == '__main__':
    # Открываем файл xlsx
    wb = openpyxl.load_workbook('20230316.xlsx')
    # Получаем все страницы из файла
    sheets = open_sheets(nums_sheets)
    # Фильтр страниц на совпадение с второй страницей
    result = filter_sheets(sheets)

    wb.save('result.xlsx')

