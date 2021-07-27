import os
from time import sleep

DIR_PATH = 'C://Users//peshi\PycharmProjects//DocxEditor//Послужные карты, для курсантов//1 взвод'

if __name__ == '__main__':
    file_names = []
    for f in os.walk(DIR_PATH + '//Служебные характеристики'):
        file_names = f[2]
    for file_path in file_names:
        if '.docx' in file_path:
            os.startfile(f'{DIR_PATH}//Служебные характеристики//{file_path}', 'print')
            sleep(5)
            os.startfile(f'{DIR_PATH}//Служебные характеристики//{file_path}', 'print')
            sleep(7)
    file_names = []
    for f in os.walk(DIR_PATH + '//Справки'):
        file_names = f[2]
    for file_path in file_names:
        if '.docx' in file_path:
            os.startfile(f'{DIR_PATH}//Справки//{file_path}', 'print')
            sleep(5)
            os.startfile(f'{DIR_PATH}//Справки//{file_path}', 'print')
            sleep(5)
            os.startfile(f'{DIR_PATH}//Справки//{file_path}', 'print')
            sleep(7)
