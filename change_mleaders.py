# Program allow to change position numbers in multileaders on a draft

import win32com.client as wc

acad_app = wc.Dispatch("AutoCAD.Application").ActiveDocument

# Define from which position start to increase/decrease values (include)
from_position = 150
# Define increment (+/-)
inc = 1

for entity in acad_app.ModelSpace:
    if entity.EntityName == 'AcDbMLeader' and entity.Visible is True:
        print(entity.TextString, end=' -> ')
        result = []
        initial_positions = entity.TextString.split('\P')

        for i in initial_positions:
            if '-' in i:
                tmp = i.split('-')
                tmp_inc = []
                for j in tmp:
                    if int(j) >= from_position:
                        j = str(int(j) + inc)
                        tmp_inc.append(j)
                    else:
                        tmp_inc.append(j)
                modified_positions = '-'.join(tmp_inc)
                result.append(modified_positions)
            elif int(i) >= from_position:
                result.append(str(int(i) + inc))
            else:
                result.append(i)

        output = '\P'.join(result)
        print(output)
        entity.TextString = output
