import openpyxl

wb              = openpyxl.load_workbook('20230316.xlsx')
sheet_bryansk   = wb.worksheets[1]
sheet_abdc      = wb.worksheets[2]

rows_bryansk    = 2
rows_abdc       = 2

max_row_bryansk = sheet_bryansk.max_row - 1
max_row_abdc    = sheet_abdc.max_row - 1

for row_abdc in range(max_row_abdc):
    A_abdc          = 'A' + str(rows_abdc)
    surname_abdc    = sheet_abdc[A_abdc].value
    print('abdc', surname_abdc)

    for row_bryansk in range(max_row_bryansk):
        A_bryansk       = 'A' + str(rows_bryansk)
        surname_bryansk = sheet_bryansk[A_bryansk].value
        print('bryansk', surname_bryansk)

        if surname_abdc == surname_bryansk:
            print(f'Совпадение: {rows_bryansk}', surname_bryansk)
        else:
            for cell in sheet_abdc[rows_abdc]:
                cell.value = None

        rows_bryansk = rows_bryansk + 1

    rows_bryansk = 2
    rows_abdc = rows_abdc + 1


wb.save('result.xlsx')