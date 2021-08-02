import os
from time import sleep
import json

try:
    with open('config.json', 'r', encoding='utf-8') as settings_file:
        settings = json.loads(settings_file.read())
except FileNotFoundError:
    print('''Не найден файл настроек config.json
Он должен находиться рядом с файлом print_files.exe (print_files.py)''')
    sleep(2)
    exit(-1)

try:
    DIR_PATH = settings['DIR_PATH']
except KeyError:
    print('Неверно указан файл настроек config.json')
    sleep(2)
    exit(-1)

if __name__ == '__main__':
    doc_to_print = None
    count_to_print = 0
    group = 0
    try:
        doc_to_print = int(input('''Какие документы вы хотите распечатать?
1 - Служебные характеристики
2 - Справки
3 - Аттестационные листы
Документ на печать: '''))
        count_to_print = int(input('''Количество экземпляров на печать: '''))
        group = int(input('''Выберите номер взвода(пример: 1): '''))
    except ValueError:
        print("Ошибка ввода данных")
        sleep(2)
        exit(-1)
    file_names = []
    path_to_walk = None
    if doc_to_print == 1:
        path_to_walk = DIR_PATH + f'\\{group} взвод\\Служебные характеристики'
    elif doc_to_print == 2:
        path_to_walk = DIR_PATH + f'\\{group} взвод\\Справки'
    elif doc_to_print == 3:
        path_to_walk = DIR_PATH + f'\\{group} взвод\\Аттестационные листы'
    else:
        print("Ошибка выбора типа документов")
        sleep(2)
        exit(-1)
    for f in os.walk(path_to_walk):
        file_names = f[2]
    for file_path in file_names:
        if '.docx' in file_path:
            for i in range(count_to_print):
                print(f'Отправляю на печать файл {file_path}')
                os.startfile(f'{path_to_walk}\\{file_path}', 'print')
                sleep(7)

    print('Все файлы были отправлены на печать')
    sleep(5)
