import calendar  

def Get_first_last_day_of_week(year, month):
    week_list = calendar.monthcalendar(year, month)
    number_of_weeks = len(week_list)
    newlist = []

    for week in week_list:

        week_a = filter((0).__ne__, week)
        week_a = list(week_a)
        newlist.append(week_a)
    first_last_day_of_week = []
    for week in newlist:
        first_last_day_of_week.append(week[:1]+week[-1:])

    return first_last_day_of_week