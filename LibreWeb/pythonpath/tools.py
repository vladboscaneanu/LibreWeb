# This is a part of LibreWeb project

def dir_as_string(argument):
    string=""
    for i in dir(argument):
        string = string + i + " -- "
    return string


