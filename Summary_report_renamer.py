
##TODO the script builds next data:

##TODO Not sure if it's needed is 4 weeks in the next month the latest week will be named as reserved and possibly hidden

##TODO "smart" Delete data from the sheets

##TODO optimize: if to run not from console
##TODO file paths input and output.




import sys
import Get_current_report
import Data_from_calendar
import Excel_file_manager
import Check_file_format

try:
    current_file_path = sys.argv[1]
except IndexError as v:
    print('Please choose a file to rename\n', v, '\n', 'Stopping script... have a good day!' )
    sys.exit(0)
Check_file_format.check_file_endwish(current_file_path)

full_file_path = Get_current_report.Get_full_file_path(current_file_path)

Check_file_format.check_file_endwish(full_file_path)

current_month_name = Get_current_report.Get_current_month_name(current_file_path)

next_month_name, next_month_name_short, next_month_index = Get_current_report.Get_next_month_name(current_file_path)

next_year, next_year_short, current_year = Get_current_report.Get_next_year(current_file_path, next_month_name)

next_month_number = next_month_index + 1
list_of_first_end_day_in_week = Data_from_calendar.Get_first_last_day_of_week(next_year,next_month_number)


print("Current month name", current_month_name)
print(next_year, next_year_short)
print(next_month_name, next_month_name_short, next_month_number)
print(list_of_first_end_day_in_week)

next_file_path = Excel_file_manager.next_file_name(current_file_path=current_file_path, current_month_name=current_month_name, next_month_name=next_month_name, current_year=current_year, next_year=next_year)
Excel_file_manager.Run_Excel_changes(full_file_path, list_of_first_end_day_in_week, next_year, next_month_name, next_month_name_short, next_year_short, next_file_path)