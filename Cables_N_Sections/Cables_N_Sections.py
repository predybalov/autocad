"""
1. Считывает данные из блока "Сечение_ПК": сечение и номера кабелей в нём
2. Форматирует и выводит в Excel таблицу сечений для ПК
3. Форматирует и выводит в Excel таблицу сечений для ТЭ6
4. Форматирует ТЭ6 согласно количеству строк в графе "Трасса прокладки кабелей" для каждого кабеля
5. Вставляет сечения в ТЭ6
6. Форматирует ТЭ6 с учётом количества строк на каждом листе
"""

import win32com.client as wc
import copy
import time

acad_app = wc.Dispatch("AutoCAD.Application").ActiveDocument.ModelSpace
xl_app = wc.Dispatch("Excel.Application").Sheets("ПКt")

start = time.time()

section_set, cable_set_unsort = [], []
cabline_1, cabline_2 = '', ''


# INPUT DATA
num_prefix = '1-'


# Iterate trough all objects (entities) in the currently opened drawing
for block in acad_app:
    if block.EntityName == 'AcDbBlockReference' and block.EffectiveName == 'Сечение_ПК' \
            and block.Visible is True and block.Layer == '0_373ПС81_Сечения_ПК':
        for attribute in block.GetAttributes():
            if attribute.TagString == 'SECTION':
                section_set.append(int(attribute.TextString[1:]))
            elif attribute.TagString == 'CABLES':
                cabline_1 = attribute.TextString
            elif attribute.TagString == 'CABLES2':
                if len(attribute.TextString) > 0:
                    cabline_2 = cabline_1 + ',' + attribute.TextString
                    cable_set_unsort.append(cabline_2.split(','))
                else:
                    cabline_2 = cabline_1
                    cable_set_unsort.append(cabline_2.split(','))


# Convert cables to float and sort
cable_set = []

for s in cable_set_unsort:
    temp_set = []
    for c in s:
        if c[-1].lower() == 'a' or c[-1].lower() == 'а':
            temp_set.append(float(c[:-1] + '.1'))
        else:
            temp_set.append(float(c))
        temp_set.sort()
    cable_set.append(temp_set)


# Create sorted dictionary {section: [cable numbers]}
common_set_temp = []
common_set_unsort, common_set = {}, {}

for i in range(len(section_set)):
    common_set_temp.append(section_set[i])
    common_set_temp.append(cable_set[i])

for i in range(0, len(common_set_temp), 2):
    common_set_unsort[common_set_temp[i]] = common_set_temp[i + 1]

cs = list(common_set_unsort.keys())
cs.sort()

for c in cs:
    common_set[c] = common_set_unsort[c]


# Create two copies of the common set for further processing
common_set_2 = copy.deepcopy(common_set)
common_set_3 = copy.deepcopy(common_set)


# Set text format for cells
xl_app.Range("R:Z").NumberFormat = "@"


# Iterate through key:value pairs in the dictionary and fill sections

# Add header
xl_app.Range("Y1").Value, xl_app.Range("Z1").Value = 'Номер сечения', 'Номер кабеля'

line = ''
length = 50
counter = 2
column = 25
i = 1

for s in common_set:
    xl_app.Cells(counter, column).Value = 'A' + str(s)
    while len(common_set[s]) > 0:
        if (len(line) + len(str(common_set[s][0]))) < length:
            if str(common_set[s][0])[-1] == '0':
                line = line + num_prefix + str(common_set[s][0])[:-2] + ', '
                common_set[s].pop(0)
            else:
                line = line + num_prefix + str(common_set[s][0])[:-2] + 'а' + ', '
                common_set[s].pop(0)
        else:
            xl_app.Cells(counter, column+1).Value = line[:-1]
            line = ''
            counter += 1

    xl_app.Cells(counter, column+1).Value = line[:-2]
    line = ''
    counter += 1
    i += 1


# Create Create sorted dictionary {cable: [section numbers]}
numbers = []
cabsec = {}

for c in common_set_2:
    numbers += common_set_2[c]
numbers_set = sorted(set(numbers))

for c in numbers_set:
    tmp = []
    for value in common_set_3:
        if c in common_set_3[value]:
            tmp.append(value)
    cabsec[c] = tmp


# Iterate through key:value pairs in the dictionary and fill cables
line = ''
counter = 1
length = 45
column = 18

for c in cabsec:
    if str(c)[-2:] == '.1':
        xl_app.Cells(counter, column).Value = num_prefix + str(c)[:-2] + 'а'
    elif str(c)[-2:] == '.0':
        xl_app.Cells(counter, column).Value = num_prefix + str(c)[:-2]
    else:
        print('Ошибка')

    space_counter = 0

    while len(cabsec[c]) > 0:
        if (len(line) + len(str(cabsec[c][0]))) < length:
            if str(cabsec[c][0])[-1] == '0':
                line = line + 'А' + str(cabsec[c][0]) + ', '
                cabsec[c].pop(0)
            else:
                line = line + 'А' + str(cabsec[c][0]) + ', '
                cabsec[c].pop(0)
        else:
            xl_app.Cells(counter, column + 1).Value = line[:-1]
            line = ''
            counter += 1
            space_counter += 1

    xl_app.Cells(counter, column + 1).Value = line[:-2]
    line = ''
    counter += 1
    i += 1
    if space_counter == 0:
        xl_app.Cells(counter, column + 1).Value = ''
        counter += 1

end_line = 2000

# Fill cable journal
for i in range(1, end_line):
    if xl_app.Cells(i, 1).Value == '' or xl_app.Cells(i, 1).Value == xl_app.Cells(i, 18).Value:
        continue
    else:
        r = 'A' + str(i) + ':' + 'P' + str(i)
        xl_app.Range(r).Insert()

for i in range(1, end_line):
    xl_app.Cells(i, 12).Value = xl_app.Cells(i, 19).Value

xl_app.Columns(18).Clear()
xl_app.Columns(19).Clear()


# Add blank lines to fit document format (first page - 24 lines, subsequent pages - 30 lines)
prev = 0
page = 1
lines1 = 24

for current in range(1, lines1 + 10):
    if xl_app.Cells(current, 1).Value != '' and xl_app.Cells(current, 1).Value is not None:
        if prev <= lines1*page and current > lines1*page + 1:
            j = prev
            step = lines1*page - prev + 1
            for j in range(step):
                r = 'A' + str(prev) + ':' + 'P' + str(prev)
                xl_app.Range(r).Insert()
            page += 1
        elif current > lines1*page:
            page += 1
        prev = current


lines2 = 30
page = 1

for current in range(25, end_line):
    if xl_app.Cells(current, 1).Value != '' and xl_app.Cells(current, 1).Value is not None:
        if prev <= lines2*page + lines1 and current > lines2*page + lines1 + 1:
            j = prev
            step = lines2*page + lines1 - prev + 1
            for j in range(step):
                r = 'A' + str(prev) + ':' + 'P' + str(prev)
                xl_app.Range(r).Insert()
            page += 1
        elif current > lines2*page + lines1:
            page += 1
        prev = current


stop = time.time()
print(round(stop - start), 'sec')
