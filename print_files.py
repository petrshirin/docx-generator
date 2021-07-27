import os
from time import sleep

DIR_PATH = './/Послужные карты, для курсантов//1 взвод'

if __name__ == '__main__':
    file_names = []
    for f in os.walk(DIR_PATH + '//Служебные характеристики'):
        file_names = f[2]
    for file_path in file_names:
        os.startfile(f'{DIR_PATH}//{file_path}', 'print')
        sleep(3)
        os.startfile(f'{DIR_PATH}//{file_path}', 'print')
        sleep(5)
    file_names = []
    for f in os.walk(DIR_PATH + '//Справки'):
        file_names = f[2]
    for file_path in file_names:
        os.startfile(f'{DIR_PATH}//{file_path}', 'print')
        sleep(3)
        os.startfile(f'{DIR_PATH}//{file_path}', 'print')
        sleep(3)
        os.startfile(f'{DIR_PATH}//{file_path}', 'print')

        sleep(5)





