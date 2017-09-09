import openpyxl
wb = openpyxl.load_workbook("Luxoft_Timesheet_27_2017_Denys_LebedevV70.xlsx", data_only=True)


ws = wb.active
alph = 'EFGHIJKL'
digit = '5'
for letter in alph:

    cell = letter + digit

    value = str(ws[cell].value)
    value = value[8:10]

    if value == '01':
        index = alph.index(letter)

        letter = alph[index]
        print(letter)

