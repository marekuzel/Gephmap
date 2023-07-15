import os

def createFileList ():
    xlsx_list = []
    current_directory = os.getcwd()
    file_list = os.listdir(current_directory)
    for file_name in file_list:
        if os.path.isfile(file_name) and file_name.endswith('.xlsx'):
            xlsx_list.append(file_name)
    return xlsx_list