import win32com.client as wc
import collections

acad_app = wc.Dispatch("AutoCAD.Application").ActiveDocument
xl_app = wc.Dispatch("Excel.Application").Sheets("Work")

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

# print(element)
# print(ppp)

xxx = []
for value in ppp:
    zzz = []
    for c in value:
        if c[-1].lower() == 'a' or c[-1].lower() == 'а':
            zzz.append(float(c[:-1] + '.1'))
        else:
            zzz.append(float(c))
        # print(c)
        zzz.sort()
    xxx.append(zzz)

# print(element)
# print(xxx)

sum = []

for i in range(len(element)):
    sum.append(element[i])
    sum.append(xxx[i])

    print(element[i])
    print(xxx[i])

print(sum)

for i in range(0, len(sum), 2):
    data[sum[i]] = sum[i + 1]

print(data)

data = collections.OrderedDict(sorted(data.items()))

print(data)


# First row is for the header, therefore the filling starts from the second row
counter = 2

# Iterate through key:value pairs in the dictionary and fill sections
for key in data:
    xl_app.Cells(counter, 1).Value = 'A' + str(key)
    tmp3 = []
    for k in data[key]:
        if str(k)[-2:] == '.1':
            tmp4 = '3‐' + str(k)[:-2] + 'a'
            tmp3.append(tmp4)
        elif str(k)[-2:] == '.0':
            tmp4 = '3‐' + str(k)[:-2]
            tmp3.append(tmp4)
        else:
            tmp4 = '3‐' + str(k)
            tmp3.append(tmp4)
    k3 = '\'' + ', '.join(tmp3)
    xl_app.Cells(counter, 2).Value = k3
    counter += 1

# Create list of the cables
numbers = []
for key in data:
    numbers += data[key]
numbers_set = sorted(set(numbers))

final_numbers = []
for cable in numbers_set:
    if cable != int(cable):
        final_numbers.append('3‐' + str(cable)[:-2] + 'a')
    else:
        final_numbers.append('3‐' + str(int(cable)))


# Iterate through key:value pairs in the dictionary and fill cables
i = 1

for cable in final_numbers:
    xl_app.Cells(i + 1, 8).Value = cable
    tmp5 = []
    for key in data:
        for value in data[key]:
            if str(value)[-2:] == '.1' and cable[2:-1] == str(value)[:-2]:
                tmp5.append('A' + str(key))
            elif str(float(value))[-2:] == '.0' and cable[2:] == str(float(value))[:-2]:
                tmp5.append('A' + str(key))
    tmp6 = ', '.join(tmp5)
    xl_app.Cells(i + 1, 9).Value = tmp6
    i += 1
