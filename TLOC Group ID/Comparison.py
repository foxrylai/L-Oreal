import os

path = input('please input log files path \n')

files = os.listdir(path)

for file in files:
    bfd_list = []
    file_path = os.path.join(path, file)
    up_counts = 0
    down_counts = 0
    with open(file_path, 'r', encoding='utf-8') as f:
        text_rows = f.readlines()
        for row in text_rows:
            if row.startswith('10.'):
                row_list = row.split()
                if row_list[2] == 'up':
                    up_counts += 1
                elif row_list[2] == 'down':
                    down_counts +=1
                bfd_list.append(row_list)
    print(f'{file.split(".")[0]} have {len(bfd_list)} bfd sessions, {up_counts} is up, {down_counts} is down')
