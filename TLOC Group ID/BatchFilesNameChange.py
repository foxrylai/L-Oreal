import os
path = input('please input log files path \n')

for file in os.listdir(path):
    if 'postlog' in file:
        new_file_temp = file.replace('-', ' ')
        new_file= new_file_temp.replace('postlog', '- after')
    else:
        new_file_temp = file.replace('-', ' ')
        new_file = new_file_temp.replace('.txt', ' - before.txt')
    old_file_path = path + '\\' + file
    new_file_path = path + '\\' + new_file
    os.renames(old_file_path,new_file_path)