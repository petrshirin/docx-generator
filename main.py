import docx
from typing import Union
import csv

DIR_PATH = ".\\Послужные карты, для курсантов\\"
TEMPLATE_ROUTES = {
    'послужные карты': ['Шаблон.docx'],
    'справки': ['.\\Шаблоны рядового\\Справка рядового.docx', '.\\Шаблоны сержанта\\Справка сержанта.docx'],
}

NUMBER_MARK_TO_WORD = {
    "3": 'удовлетворительно',
    "4": 'хорошо',
    "5": 'отлично',
}

DATA_FILES = ["1_DATA.csv", "2_DATA.csv", "3_DATA.csv"]
DIR_TO_SAVE = "взвод"
CERTIFICATION_LIST_TEMPLATES = {
    'рядовой': {
        'удовлетворительно': '.\\Шаблоны рядового\\Аттестационный лист рядового удовлетворительно.docx',
        'хорошо': '.\\Шаблоны рядового\\Аттестационный лист рядового хорошо.docx',
        'отлично': '.\\Шаблоны рядового\\Аттестационный лист рядового отлично.docx',
    },
    'сержант': {
        'удовлетворительно': '.\\Шаблоны сержанта\\Аттестационный лист сержанта удовлетворительно.docx',
        'хорошо': '.\\Шаблоны сержанта\\Аттестационный лист сержанта хорошо.docx',
        'отлично': '.\\Шаблоны сержанта\\Аттестационный лист сержанта отлично.docx',
    }
}
SERVICE_CHARACTERISTIC_TEMPLATES = {
    'рядовой': {
        'удовлетворительно': '.\\Шаблоны рядового\\Служебная характеристика рядового удовлетворительно.docx',
        'хорошо': '.\\Шаблоны рядового\\Служебная характеристика рядового хорошо.docx',
        'отлично': '.\\Шаблоны рядового\\Служебная характеристика рядового отлично.docx',
    },
    'сержант': {
        'удовлетворительно': '.\\Шаблоны сержанта\\Служебная характеристика сержанта удовлетворительно.docx',
        'хорошо': '.\\Шаблоны сержанта\\Служебная характеристика сержанта хорошо.docx',
        'отлично': '.\\Шаблоны сержанта\\Служебная характеристика сержанта отлично.docx',
    }
}


def create_docx_file_by_template(dir_path: str, template_name: str) -> docx.Document:
    docx_file = docx.Document(dir_path + template_name)
    return docx_file


def update_cell_value(self: docx.document.Document,
                      table_id: int,
                      row_id: int,
                      col_id: int,
                      value: Union[str, list],
                      method: int = 1) -> None:
    """

    :param self:
    :param table_id:
    :param row_id:
    :param col_id:
    :param value:
    :param method:
    1 - ADD IN CELL,
    2 - REPLACE ALL CELL VALUE,
    3 - REPLACE '_' SYMBOLS,
    4 - DELETE VALUE,
    5 - DELETE VALUE BY STR
    :return:
    """
    paragraph = self.tables[table_id].rows[row_id].cells[col_id].paragraphs[-1]
    if method == 1:
        r = paragraph.add_run(value)
        r.italic = True
    elif method == 2:
        self.tables[table_id].rows[row_id].cells[col_id].text = ""
        r = paragraph.add_run(value)
        r.italic = True
    elif method == 3:
        self.tables[table_id].rows[row_id].cells[col_id].text = self.tables[table_id].rows[row_id].cells[
            col_id].text.replace('_', '')
        r = paragraph.add_run(value)
        r.italic = True
    elif method == 4:
        self.tables[table_id].rows[row_id].cells[col_id].text = ""
    elif method == 5:
        self.tables[table_id].rows[row_id].cells[col_id].text = self.tables[table_id].rows[row_id].cells[
            col_id].text.replace(value, '')
    elif method == 6:
        i = 0
        for run in paragraph.runs:
            if run.text == '{}':
                run.text = value[i]
                run.italic = True
                i += 1


def show_template_cols():
    docx_file = docx.Document(DIR_PATH + TEMPLATE_ROUTES['послужные карты'][0])
    i = 0
    for table in docx_file.tables:
        j = 0
        for row in table.rows:
            k = 0
            for cell in row.cells:
                g = 0
                for par in cell.paragraphs:
                    print(i, j, k, g, [run.text for run in par.runs])
                    g += 1
                k += 1
            j += 1
        i += 1


def show_template_paragraphs():
    docx_file = docx.Document(DIR_PATH + SERVICE_CHARACTERISTIC_TEMPLATES['сержант']['3'])
    i = 0
    for par in docx_file.paragraphs:
        print(i, par.text)
        i += 1


def create_posluj_cards(template: str, data_file: str = DATA_FILES[0], dir_to_save: str = DIR_TO_SAVE):
    with open(data_file, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f, delimiter='|')
        for row in reader:
            docx_file = create_docx_file_by_template(
                DIR_PATH,
                template)
            update_cell_value(docx_file, 0, 1, 0, row['ФИО'], 1)
            update_cell_value(docx_file, 0, 2, 0, [row['Дата рождения']], 6)
            update_cell_value(docx_file, 0, 3, 0, row['Место рождения'], 1)
            update_cell_value(docx_file, 0, 4, 0, 'русский', 5)
            update_cell_value(docx_file, 0, 5, 1,
                              f'ФГБОУ ВО "Иркутский государственный университет", {row["Год_окончания_учебы"]}', 1)
            update_cell_value(docx_file, 0, 7, 2, [row['Серия'], row['Номер'], row['Дата выдачи'], row['Кем выдан']], 6)
            update_cell_value(docx_file, 0, 12, 1, row['ИНН'], 1)
            update_cell_value(docx_file, 1, 0, 1, row['ВК'], 1)
            update_cell_value(docx_file, 1, 1, 1, row['Адрес регистрации'], 1)
            update_cell_value(docx_file, 1, 2, 1, row['Фактический адрес'], 1)
            update_cell_value(docx_file, 2, 0, 0, row['ФИО'], 1)
            update_cell_value(docx_file, 2, 2, 0, row['Дата рождения'], 1)
            update_cell_value(docx_file, 2, 7, 0, [row['Рост'],
                                                   row['Голова'],
                                                   row['Обувь'],
                                                   row['Противогаз'],
                                                   row['Оммуниция']], 6)
            update_cell_value(docx_file, 2, 8, 1, row['Фактический адрес'], 1)
            docx_file.save(DIR_PATH + f"{dir_to_save}\\Послужные карты\\Послужная карта {row['ФИО']}.docx")


def create_help_cards(template: str, data_file: str = DATA_FILES[0], dir_to_save: str = DIR_TO_SAVE):
    with open(data_file, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f, delimiter='|')
        for row in reader:
            docx_file = create_docx_file_by_template(
                DIR_PATH,
                template)
            values = [row['ФИОдат'], row['ВТП'], row['Итог']]
            i = 0
            for paragraph in docx_file.paragraphs:
                if i > len(values):
                    break
                try:
                    for run in paragraph.runs:
                        if run.text.strip() == '{}':
                            run.text = values[i]
                            i += 1
                            run.underline = True
                except Exception as e:
                    print(values)
                    raise e
            docx_file.save(DIR_PATH + f"{dir_to_save}\\Справки\\Справка {row['ФИО']}.docx")


def create_service_characteristics(template: dict, data_file: str = DATA_FILES[0], dir_to_save: str = DIR_TO_SAVE):
    with open(data_file, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f, delimiter='|')
        for row in reader:
            try:
                docx_file = create_docx_file_by_template(
                    DIR_PATH,
                    template[row['Итог']])
            except KeyError:
                continue
            values = [row['ФИОрод'], row['ФИО']]
            i = 0
            for paragraph in docx_file.paragraphs:
                if i > len(values):
                    break
                try:
                    for run in paragraph.runs:
                        if run.text.strip() == '{}':
                            run.text = values[i]
                            i += 1
                            run.underline = True
                except Exception as e:
                    print(values)
                    raise e
            docx_file.save(
                DIR_PATH + f"{dir_to_save}\\Служебные характеристики\\Служебная характеристика {row['ФИО']}.docx")


def create_certifications_lists(template: dict, data_file: str = DATA_FILES[0], dir_to_save: str = DIR_TO_SAVE):
    with open(data_file, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f, delimiter='|')
        for row in reader:
            docx_file = create_docx_file_by_template(
                DIR_PATH,
                template.get(row['Итог'], template['удовлетворительно']))
            values = [row['ФИОдат'],
                      row['Дата рождения'],
                      row['Место рождения'],
                      row['Факультет'],
                      row['Год_окончания_учебы'], 
                      row['Гр_специальность']]
            i = 0
            for paragraph in docx_file.paragraphs:
                if i > len(values):
                    break
                try:
                    for run in paragraph.runs:
                        if run.text.strip() == '{}':
                            run.text = values[i]
                            i += 1
                            run.underline = True
                except Exception as e:
                    print(values)
                    raise e
            docx_file.save(DIR_PATH + f"{dir_to_save}\\Аттестационные листы\\Аттестационный лист {row['ФИО']}.docx")


if __name__ == '__main__':
    create_posluj_cards(TEMPLATE_ROUTES['послужные карты'][0], DATA_FILES[0], '1 взвод')
    create_posluj_cards(TEMPLATE_ROUTES['послужные карты'][0], DATA_FILES[1], '2 взвод')
    create_posluj_cards(TEMPLATE_ROUTES['послужные карты'][0], DATA_FILES[2], '3 взвод')

    create_help_cards(TEMPLATE_ROUTES['справки'][1], DATA_FILES[0], '1 взвод')
    create_help_cards(TEMPLATE_ROUTES['справки'][0], DATA_FILES[1], '2 взвод')
    create_help_cards(TEMPLATE_ROUTES['справки'][0], DATA_FILES[2], '3 взвод')

    create_certifications_lists(CERTIFICATION_LIST_TEMPLATES['сержант'], DATA_FILES[0], '1 взвод')
    create_certifications_lists(CERTIFICATION_LIST_TEMPLATES['рядовой'], DATA_FILES[1], '2 взвод')
    create_certifications_lists(CERTIFICATION_LIST_TEMPLATES['рядовой'], DATA_FILES[2], '3 взвод')

    create_service_characteristics(SERVICE_CHARACTERISTIC_TEMPLATES['сержант'], DATA_FILES[0], '1 взвод')
    create_service_characteristics(SERVICE_CHARACTERISTIC_TEMPLATES['рядовой'], DATA_FILES[1], '2 взвод')
    create_service_characteristics(SERVICE_CHARACTERISTIC_TEMPLATES['рядовой'], DATA_FILES[2], '3 взвод')
