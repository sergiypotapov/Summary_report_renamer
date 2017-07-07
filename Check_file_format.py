import sys

def check_file_endwish(current_file_path):
    current_file_path = current_file_path

    if current_file_path.endswith('xlsx'):
        print ("xlsx OK")
    elif current_file_path.startswith('L'):
        print("L ok")
    else:
        print("Looks it's wrong file or file path:\n", current_file_path, '\n  Stopping script... have a good day!')
        sys.exit(0)