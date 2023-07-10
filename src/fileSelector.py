import os

# Get the current directory
current_directory = os.getcwd()

# Get a list of all files and directories in the current directory
file_list = os.listdir(current_directory)

# Iterate through the file list and print only the files
for file_name in file_list:
    if os.path.isfile(file_name):
        print(file_name)
