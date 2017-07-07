import check_if_hour_in_month
import string

def get_hours_for_code(read_code_letter, read_code_digit, ws):

    # check if hour in needed month

    check_if_hour_in_month.check_if_hours_in_month()

    # index of E letter
    index = 4
    code_total = 0
    ifint = True



    read_code_letter = string.ascii_uppercase[index]
    stop_code = 'L'+ str(read_code_digit)

    read_code =  read_code_letter + str( read_code_digit)



    while read_code != stop_code:



        read_code_value = ws[read_code].value



        if ifint == isinstance(read_code_value, int):

            code_total = code_total + read_code_value

            index = index + 1

            read_code_letter = string.ascii_uppercase[index]
            read_code = read_code_letter + str(read_code_digit)



        else:
            index = index + 1
            read_code_letter = string.ascii_uppercase[index]
            read_code = read_code_letter + str( read_code_digit)


    return(code_total)
    print("horrey")