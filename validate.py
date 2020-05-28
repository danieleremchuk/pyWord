from pymsgbox import alert

def valid_file(my_file):
    if len(my_file) == 0:
        return False
    if ".doc" not in my_file:
        return False
    else:
        return True
    
def valid_file_msg(file_valid):
    if not valid_file(file_valid):
        return "You must select a file first before proceeding."

def valid_name(my_name):
    if len(my_name) == 0:
        return False
    else:
        return True

def valid_name_msg(name_valid):
    if not valid_name(name_valid):
        return "You must enter a brewery name before proceeding."

    
