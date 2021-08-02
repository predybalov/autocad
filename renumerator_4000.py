import win32com.client
from pyautocad import Autocad, APoint, ACAD
import time

def report():
    print()
    acad = Autocad(create_if_not_exists=False)
    print('*' * 50)
    if counter % 10 == 1:
        print(f'Изменён {counter} элемент за {program_runtime} с.')
    elif counter % 10 == 2 or counter % 10 == 3 or counter % 10 == 4:
        print(f'Изменёно {counter} элемента за {program_runtime} с.')
    else:
        print(f'Изменёно {counter} элементов за {program_runtime} с.')
    print('в файле:', acad.doc.Name)
    print('*' * 50)

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument  # Document object

prefix = input('Введите префикс: ')
prefix_length = len(prefix)
suffix = input('Введите суффикс: ')
suffix_length = len(suffix)
start_num = int(input('Введите первый номер диапазона: '))
finish_num = int(input('Введите последний номер диапазона: '))
increment = int(input('Введите инкремент: '))

print()
print('*' * 50)
print('Начинаю переименование...')
print('*' * 50)
print()

counter = 0

start_time = time.time()

# iterate trough all objects (entities) in the currently opened drawing
# and if its a BlockReference, display its attributes and some other things.
for entity in doc.ModelSpace:
    name = entity.EntityName
    if name == 'AcDbBlockReference':
        HasAttributes = entity.HasAttributes
        if HasAttributes:
            # print(entity.Name)
            for attrib in entity.GetAttributes():
                # first condition - latin alphabet, second - russian alphabet
                if (attrib.TextString[:(prefix_length)] == prefix and
                        start_num <= int((attrib.TextString[prefix_length:])[:-suffix_length]) <= finish_num):
                    current_num = int((attrib.TextString[prefix_length:])[:-suffix_length])
                    modified_num = str(current_num + increment)
                    modText = prefix + modified_num + suffix
                    print(attrib.TextString, '->', modText)
                    attrib.TextString = modText
                    counter += 1

end_time = time.time()
program_runtime = int(end_time - start_time)

report()








