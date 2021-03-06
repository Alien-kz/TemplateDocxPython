# Установка
```
python3 -m pip install python-docx pandas docx2txt
```

# Использование
Необходимо создать шаблон template.docx, используя имена подстановок. Например:
```
Пример NUMBER для шаблона.
Сохранение форматирования NAME.
```

Создать таблицу со значениями подстановок data.csv (можно создать в excel), где в первой строке имена подстановок, в последующих строках их значения. Например:
```
FILE,NUMBER,NAME
название_файла_1,1,Ирина
название_файла_2,2,Кайрат
другое_название,2022,Монти
```

# Прямая задача (по шаблону и таблице подстановок сгенерировать пачку файлов)
Из файлов data.csv и template.docx генерируются файлы .docx и сохраняются в директорию output/:
```
python3 fill.py
```

# Обратная задача (по пачке файлов сгенерировать таблицу подстановок)
В папке data/ перебираем все файлы docx и вытаскиваем текст от "КОМИСИИ" до "Присутствовали" и сохраняем в файл out.csv
```
python3 extract.py data/ "КОМИССИИ" "Присутствовали" out.csv
```

Отладка
```
python3 extract.py data/ "КОМИСИИ" "Присутствовали" out.csv --debug 1
```

Если в многострочном результате нужна только одна строка, например вторая с конца, можно указать необязательным аргументом (--shift -2).
```
python3 extract.py data/ "КОМИССИИ" "Присутствовали" HHMM2.csv --shift -2
```
