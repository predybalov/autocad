"""
1. Считывает данные из блока "Сечение_ПК": сечение и номера кабелей в нём
2. Форматирует и выводит в Excel таблицу сечений для ПК
3. Форматирует и выводит в Excel таблицу сечений для ТЭ6
4. Форматирует ТЭ6 согласно количеству строк в графе "Трасса прокладки кабелей" для каждого кабеля
5. Вставляет сечения в ТЭ6
6. Форматирует ТЭ6 с учётом количества строк на каждом листе
"""

import win32com.client as wc
import collections
import copy
import time

acad_app = wc.Dispatch("AutoCAD.Application").ActiveDocument.ModelSpace
xl_app = wc.Dispatch("Excel.Application").Sheets("Work")

start = time.time()

data_set = {}
section_set = []
cable_set = []
cabline_1, cabline_2 = '', ''

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
                    cable_set.append(cabline_2.split(','))
                else:
                    cabline_2 = cabline_1
                    cable_set.append(cabline_2.split(','))








xxx = []
for value in cable_set:
    zzz = []
    for c in value:
        if c[-1].lower() == 'a' or c[-1].lower() == 'а':
            zzz.append(float(c[:-1] + '.1'))
        else:
            zzz.append(float(c))
        zzz.sort()
    xxx.append(zzz)



sum = []

for i in range(len(section_set)):
    sum.append(section_set[i])
    sum.append(xxx[i])

for i in range(0, len(sum), 2):
    data_set[sum[i]] = sum[i + 1]


data_set = collections.OrderedDict(sorted(data_set.items()))


###
data2 = copy.deepcopy(data_set)
###
data3 = copy.deepcopy(data_set)
###
#
#
# # Set text format for cells
# xl_app.Range("R:Z").NumberFormat = "@"
#
#
# # Iterate through key:value pairs in the dictionary and fill sections
# # Add header
# xl_app.Range("Y1").Value, xl_app.Range("Z1").Value = 'Номер сечения', 'Номер кабеля'
#
# line = ''
# length = 50
# counter = 2
# column = 25
# i = 1
#
# for key in section_set:
#     xl_app.Cells(counter, column).Value = 'A' + str(key)
#     while len(section_set[key]) > 0:
#         if (len(line) + len(str(section_set[key][0]))) < length:
#             if str(section_set[key][0])[-1] == '0':
#                 line = line + '3-' + str(section_set[key][0])[:-2] + ', '
#                 section_set[key].pop(0)
#             else:
#                 line = line + '3-' + str(section_set[key][0])[:-2] + 'а' + ', '
#                 section_set[key].pop(0)
#         else:
#             xl_app.Cells(counter, column+1).Value = line[:-1]
#             line = ''
#             counter += 1
#
#     xl_app.Cells(counter, column+1).Value = line[:-2]
#     line = ''
#     counter += 1
#     i += 1
#
#
# # Create list of the cables
# numbers = []
# for key in data2:
#     numbers += data2[key]
# numbers_set = sorted(set(numbers))
#
#
# final_numbers = []
# for cable in numbers_set:
#     if cable != int(cable):
#         final_numbers.append(str(cable))
#     else:
#         final_numbers.append(str(cable))
#
#
# data10 = {}
#
# for number in final_numbers:
#     tmpf = []
#     for value in data3:
#         if float(number) in data3[value]:
#             tmpf.append(value)
#
#     data10[float(number)] = tmpf
#
# data11 = collections.OrderedDict(sorted(data10.items()))
#
# # Iterate through key:value pairs in the dictionary and fill cables
# line = ''
# counter = 1
# length = 45
# column = 18
#
# for key in data11:
#     if str(key)[-2:] == '.1':
#         xl_app.Cells(counter, column).Value = '3-' + str(key)[:-2] + 'а'
#     elif str(key)[-2:] == '.0':
#         xl_app.Cells(counter, column).Value = '3-' + str(key)[:-2]
#     else:
#         print('Ошибка')
#
#     mount = 0
#
#     while len(data11[key]) > 0:
#         if (len(line) + len(str(data11[key][0]))) < length:
#             if str(data11[key][0])[-1] == '0':
#                 line = line + 'А' + str(data11[key][0]) + ', '
#                 data11[key].pop(0)
#             else:
#                 line = line + 'А' + str(data11[key][0]) + ', '
#                 data11[key].pop(0)
#         else:
#             xl_app.Cells(counter, column + 1).Value = line[:-1]
#             line = ''
#             counter += 1
#             mount += 1
#
#     xl_app.Cells(counter, column + 1).Value = line[:-2]
#     line = ''
#     counter += 1
#     i += 1
#     if mount == 0:
#         xl_app.Cells(counter, column + 1).Value = ''
#         counter += 1
#
# end_line = 2000
#
# # Fill cable journal
# for i in range(1, end_line):
#     if xl_app.Cells(i, 1).Value == '' or xl_app.Cells(i, 1).Value == xl_app.Cells(i, 18).Value:
#         continue
#     else:
#         r = 'A' + str(i) + ':' + 'P' + str(i)
#         xl_app.Range(r).Insert()
#
# for i in range(1, end_line):
#     xl_app.Cells(i, 12).Value = xl_app.Cells(i, 19).Value
#
# xl_app.Columns(18).Clear()
# xl_app.Columns(19).Clear()
#
#
# # Add blank lines to fit document format (first page - 24 lines, subsequent pages - 30 lines)
# prev = 0
# page = 1
# lines1 = 24
#
# for current in range(1, lines1 + 10):
#     if xl_app.Cells(current, 1).Value != '' and xl_app.Cells(current, 1).Value is not None:
#         if prev <= lines1*page and current > lines1*page + 1:
#             j = prev
#             step = lines1*page - prev + 1
#             for j in range(step):
#                 r = 'A' + str(prev) + ':' + 'P' + str(prev)
#                 xl_app.Range(r).Insert()
#             page += 1
#         elif current > lines1*page:
#             page += 1
#         prev = current
#
#
# lines2 = 30
# page = 1
#
# for current in range(25, end_line):
#     if xl_app.Cells(current, 1).Value != '' and xl_app.Cells(current, 1).Value is not None:
#         if prev <= lines2*page + lines1 and current > lines2*page + lines1 + 1:
#             j = prev
#             step = lines2*page + lines1 - prev + 1
#             for j in range(step):
#                 r = 'A' + str(prev) + ':' + 'P' + str(prev)
#                 xl_app.Range(r).Insert()
#             page += 1
#         elif current > lines2*page + lines1:
#             page += 1
#         prev = current
#
#
# stop = time.time()
# print(round(stop - start), 'sec')
