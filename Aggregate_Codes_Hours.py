import openpyxl
import Get_hours_for_code

def Run_Excel_changes():
    wb = openpyxl.load_workbook("Luxoft_Timesheet_27_2017_Denys_LebedevV70.xlsx", data_only=True)
    ws = wb.active

    read_code_letter = 'C'
    read_code_digit = 7
    read_code = read_code_letter + str(read_code_digit)
    read_code_value = ws[read_code].value

    dict_code_hours = {}

#read Project Codes

    while read_code_value !=0:

        read_code_value = ws[read_code].value

        if read_code_value !=0:
            code_total = Get_hours_for_code.get_hours_for_code(read_code_letter, read_code_digit, ws)

            read_code_digit = read_code_digit + 1
            read_code = read_code_letter + str (read_code_digit)

            print(read_code_value, code_total)



Run_Excel_changes()