import os
path = input('please input log files path \n')

for file in os.listdir(path):
    new_file = file.replace('  ', ' ', 1)
    new_file = new_file.lower()
    old_file_path = path + '\\' + file
    new_file_path = path + '\\' + new_file
    os.renames(old_file_path,new_file_path)