import win32com.client as wc

xl_app = wc.Dispatch("Excel.Application").Sheets("Work")

# s = '3,2,5,8,5,34,53,65,7556,87,98,23,12,34,444,333,222,111,888,777,666,555,444,333,345,348,3,2,7'
# s2 = '3,2,5,8,5,34,53,65,7556,87,98,23,12,34,444,333,222,111,888,777,666,555,444,333,345,348,3,2,7'
#
# s_split = s.split(',')
# lst = []
#
#
# for c in s_split:
#     lst.append('3-' + c)

data = {1: ['3-444', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333'],
        2: ['3-222', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333'],
        44: ['3-3', '3-2', '3-5', '3-8', '3-5', '3-34', '3-53', '3-65', '3-7556', '3-87', '3-98', '3-23', '3-12', '3-34', '3-444', '3-333', '3-222', '3-111', '3-888', '3-777', '3-666', '3-555', '3-444', '3-333', '3-345', '3-348', '3-3', '3-2', '3-7']}

two = [['3-3', '3-2', '3-5', '3-8', '3-5', '3-34', '3-53', '3-65', '3-7556', '3-87', '3-98', '3-23', '3-12', '3-34', '3-444', '3-333', '3-222', '3-111', '3-888', '3-777', '3-666', '3-555', '3-444', '3-333', '3-345', '3-348', '3-3', '3-2', '3-7'],
       ['3-111', '3-2', '3-5', '3-8', '3-5', '3-34', '3-53', '3-65', '3-7556', '3-87', '3-98', '3-23', '3-12', '3-34', '3-444', '3-333', '3-222', '3-111', '3-888', '3-777', '3-666', '3-555', '3-444', '3-333', '3-345', '3-348', '3-3', '3-2', '3-7'],
       ['3-444', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348', '3-333', '3-345', '3-348']]

line = ''
length = 50
counter = 1
i = 1

for key in data:
    # print(key)
    # print(data[key])

    print('A' + str(key))
    xl_app.Cells(counter, 1).Value = 'A' + str(key)
    while len(data[key]) > 0:
        if (len(line) + len(data[key][0])) < length:
            line = line + data[key][0] + ', '
            data[key].pop(0)
        else:
            print(line[:-1])
            xl_app.Cells(counter, 2).Value = line[:-1]
            print(len(line))
            line = ''
            counter += 1
    print(line[:-2])
    xl_app.Cells(counter, 2).Value = line[:-2]
    print(len(line))
    print()
    line = ''
    counter += 1
    i += 1