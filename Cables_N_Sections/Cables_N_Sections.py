import win32com.client as wc
import collections
import copy
import time

acad_app = wc.Dispatch("AutoCAD.Application").ActiveDocument
xl_app = wc.Dispatch("Excel.Application").Sheets("Work")

start = time.time()

data = {}
element = []
c = []
tmpc1, tmpc2 = '', ''
ppp = []

# Iterate trough all objects (entities) in the currently opened drawing
for entity in acad_app.ModelSpace:
    # Specify block Name as EffectiveName and layer name as Layer
    if entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == 'Сечение_ПК' \
            and entity.Visible == True and entity.Layer == '0_373ПС81_Сечения_ПК':
        for attrib in entity.GetAttributes():
            if attrib.TagString == 'SECTION':
                element.append(int(attrib.TextString[1:]))
            elif attrib.TagString == 'CABLES':
                tmpc1 = attrib.TextString
            elif attrib.TagString == 'CABLES2':
                if len(attrib.TextString) > 0:
                    tmpc2 = tmpc1 + ',' + attrib.TextString
                    tc2 = tmpc2.split(',')
                    ppp.append(tc2)
                else:
                    tmpc2 = tmpc1
                    tc2 = tmpc2.split(',')
                    ppp.append(tc2)

xxx = []
for value in ppp:
    zzz = []
    for c in value:
        if c[-1].lower() == 'a' or c[-1].lower() == 'а':
            zzz.append(float(c[:-1] + '.1'))
        else:
            zzz.append(float(c))
        zzz.sort()
    xxx.append(zzz)


sum = []

for i in range(len(element)):
    sum.append(element[i])
    sum.append(xxx[i])

for i in range(0, len(sum), 2):
    data[sum[i]] = sum[i + 1]


data = collections.OrderedDict(sorted(data.items()))



###
data2 = copy.deepcopy(data)
###
data3 = copy.deepcopy(data)
###


# Iterate through key:value pairs in the dictionary and fill sections
line = ''
length = 50
counter = 1
column = 25
i = 1

for key in data:
    xl_app.Cells(counter, column).Value = 'A' + str(key)
    while len(data[key]) > 0:
        if (len(line) + len(str(data[key][0]))) < length:
            if str(data[key][0])[-1] == '0':
                line = line + '3-' + str(data[key][0])[:-2] + ', '
                data[key].pop(0)
            else:
                line = line + '3-' + str(data[key][0])[:-2] + 'а' + ', '
                data[key].pop(0)
        else:
            xl_app.Cells(counter, column+1).Value = line[:-1]
            line = ''
            counter += 1

    xl_app.Cells(counter, column+1).Value = line[:-2]
    line = ''
    counter += 1
    i += 1


# Create list of the cables
numbers = []
for key in data2:
    numbers += data2[key]
numbers_set = sorted(set(numbers))


final_numbers = []
for cable in numbers_set:
    if cable != int(cable):
        final_numbers.append(str(cable))
    else:
        final_numbers.append(str(cable))


data10 = {}

for number in final_numbers:
    tmpf = []
    for value in data3:
        if float(number) in data3[value]:
            tmpf.append(value)

    data10[float(number)] = tmpf

data11 = collections.OrderedDict(sorted(data10.items()))

# Iterate through key:value pairs in the dictionary and fill cables
counter = 1
length = 45
column = 18

for key in data11:
    if str(key)[-2:] == '.1':
        xl_app.Cells(counter, column).Value = '3-' + str(key)[:-2] + 'а'
    elif str(key)[-2:] == '.0':
        xl_app.Cells(counter, column).Value = '3-' + str(key)[:-2]
    else:
        print('Ошибка')

    mount = 0

    while len(data11[key]) > 0:
        if (len(line) + len(str(data11[key][0]))) < length:
            if str(data11[key][0])[-1] == '0':
                line = line + 'А' + str(data11[key][0]) + ', '
                data11[key].pop(0)
            else:
                line = line + 'А' + str(data11[key][0]) + ', '
                data11[key].pop(0)
        else:
            xl_app.Cells(counter, column + 1).Value = line[:-1]
            line = ''
            counter += 1
            mount += 1

    xl_app.Cells(counter, column + 1).Value = line[:-2]
    line = ''
    counter += 1
    i += 1
    if mount == 0:
        xl_app.Cells(counter, 6).Value = ''
        counter += 1


# Fill cable journal
for i in range(1, 2000):
    if xl_app.Cells(i, 1).Value == '' or xl_app.Cells(i, 1).Value == xl_app.Cells(i, 18).Value:
        continue
    else:
        r = 'A' + str(i) + ':' + 'P' + str(i)
        xl_app.Range(r).Insert()

for i in range(1, 2000):
    xl_app.Cells(i, 12).Value = xl_app.Cells(i, 19).Value


xl_app.Columns(18).Clear()
xl_app.Columns(19).Clear()

stop = time.time()

print(round(stop - start))
