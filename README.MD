# Генератор документов для ВУЦ
___

### Запуск из кода
Зависимости: Python 3.x, Word версии 16 или выше(можно и более ранний, но не проверялось)
1. `git clone https://github.com/petrshirin/docx-generator.git`
2. Создаем виртуальное окружение(необязательный пункт)
3. Устанавливаем зависимости (`pip install -r req.txt`)
4. Создаем файлы 1_DATA.csv, 2_DATA.csv, 3_DATA.csv из шаблона  "Пример данных.csv"
5. Заполняем данные в csv файлах (1_DATA.csv, 2_DATA.csv, 3_DATA.csv)
6. Запускаем скрипт main.py и следуем инструкциям

### Запуск с помощью exe файлов(если ничего не менялось в коде)
1. `git clone https://github.com/petrshirin/docx-generator.git`
2. Создаем файлы 1_DATA.csv, 2_DATA.csv, 3_DATA.csv из шаблона  "Пример данных.csv"
3. Заполняем данные в csv файлах (1_DATA.csv, 2_DATA.csv, 3_DATA.csv)
4. Запускаем файл `main.exe` и следуем инструкциям

### Полуавтоатическая печать файлов
В проекте есть файлы print_files.py/print_files.exe, которая сама отправляет на печать файлы, 
человеку остается только следить за бумагой и забирать готовые документы
**Перед запуском нужно сгенерировать файлы по одной из инструкции выше!**
> Запуск печати с помощью python скрипта
> > `python print_files.py`

> Запуск с помощью exe файла
> > Запускаем файл print_files.exe

### Описание csv
Описание данных, которые нужны для работы скрипта

**Данные можно заполнять в exel, Затем, нужно созранить файл как csv с разделителем `;`**
> Пример есть в файле "Пример данных.csv"

1. ФИО - как в паспорте
2. ФИО в дательном - проверить склонение фамилии
3. ФИО в родительном - проверить склонение фамилии
4. Дата рождения - берем из паспорта пример формата `1 июля 2021`
5. Место рождения - полностью как в паспорте
6. Серия - как в паспорте 
7. Номер - как в паспорте 
8. Дата выдачи - как в паспорте `1.06.2021`
9. Кем выдан - как в паспорте 
10. Адрес регистрации - полный адрес как в паспорте 
11. Фактический адрес - полный адрес взять 
12. Рост - высота человека 
13. Голова - размер головного убора 
14. Обувь - размер обуви 
15. Противогаз - размер противогаза 
16. Оммуниция - размер летнего кителя 
17. ВК - Военный коммисариат, за котором закреплен курсант 
18. Факультет - факультет, на котором обучается, пишется полностью с большой буквы только первое слово 
19. Гражданская специальность - как в дипломе, см пример 
20. Год окончания учебы - число 
21. ИНН - ИНН 
22. ВТП - оценка за экзамен по военно тактической подготовке 
23. Итог - итоговая оценка за экзамен на сборах

### Описание файла конфигурации
В файле есть разделение на сержантов и рядовых, эту нужно для правильного выбора файла при генерации документов
а также по итоговой оценке на сборах генерятся разные характеристики и аттестационные листы
+ DIR_PATH - строка, папка для хранения документов
+ TEMPLATE_ROUTES - содержит в себе пути к шаблонам послужных карт и справок
+ NUMBER_MARK_TO_WORD переводит оценку из данных в слово
+ DATA_FILES путь до файлов с данными
+ CERTIFICATION_LIST_TEMPLATES - содержит в себе пути к шаблону Аттестационных листов
+ SERVICE_CHARACTERISTIC_TEMPLATES - содержит в себе пути к шаблону Служебнах характеристик
 