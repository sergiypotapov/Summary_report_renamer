import openpyxl
wb = openpyxl.load_workbook('tabs.xlsx', data_only=True)
ws = wb.get_sheet_by_name('a')

a4 = ws['A4'].value


a4 = str(a4)
print(a4)
wb2 = openpyxl.load_workbook('tabs.xlsx', data_only=False)
ws = wb2.get_sheet_by_name('a')
d4 = ws['D4'].value = 20

ws.title = a4
wb2.save('caramba.xlsx')
