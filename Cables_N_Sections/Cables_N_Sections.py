import win32com.client as wc
import collections

acad_app = wc.Dispatch("AutoCAD.Application").ActiveDocument
xl_app = wc.Dispatch("Excel.Application").Sheets("Work")

data = {}
element = []
c = []

# Iterate trough all objects (entities) in the currently opened drawing
for entity in acad_app.ModelSpace:
# Specify block Name as EffectiveName and layer name as Layer
    if entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == 'Сечение_ПК' \
            and entity.Visible == True and entity.Layer == '0_373ПС81_Сечения_ПК':
        for attrib in entity.GetAttributes():
            if attrib.TagString == 'SECTION':
                element.append(int(attrib.TextString[1:]))
            elif attrib.TagString == 'CABLES':
                tmp1 = []
                tmp2 = attrib.TextString.split(',')
                for value in tmp2:
                    # If the cable contains the character
                    if value[-1].lower() == 'a' or value[-1].lower() == 'а':
                        tmp1.append(float(value[:-1] + '.1'))
                    else:
                        tmp1.append(float(value))
                tmp1.sort()
                element.append(tmp1)

for i in range(0, len(element), 2):
    data[element[i]] = element[i + 1]

# Because field in attribute value limited by 256 symbols
# and some sections are not fit into the field, the variable below is used for solve this
# big_values = {0: [0,0,0],
#               1: [0,0,0]}
#
#
#
# data.update(big_values)
data = collections.OrderedDict(sorted(data.items()))

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
