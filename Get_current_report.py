#import openpyxl
import os
import calendar

arrey = calendar.month_name

List_of_months = list(arrey)
List_of_months.pop(0)

#print(List_of_months)

def Get_full_file_path(file_path):
    file_path = os.path.realpath(file_path)

    return file_path

def Get_current_month_name(file_path):
    splitted = file_path.split()
    current_month_name = set(splitted) & set(List_of_months)
    current_month_name =list(current_month_name)
    current_month_name = current_month_name[0]

    return current_month_name

def Get_next_year(file_path, next_month_name):
    splitted = file_path.split()
    next_year = splitted[-1:]
    next_year = next_year[0]
    next_year = next_year[:-5]
    next_year, current_year = int(next_year), int(next_year)
    if next_month_name == "January":

        next_year+= 1

    next_year_short = next_year - 2000
    return next_year, next_year_short, current_year

def Get_next_month_name(file_path):
    current_month_name = Get_current_month_name(file_path)
    current_index = List_of_months.index(current_month_name)

    if current_index < 11:
        next_index = current_index + 1
    else:
        next_index = 0

    next_month_name = List_of_months[next_index]
    next_month_name_short = next_month_name[:3]


    return next_month_name, next_month_name_short, next_index





