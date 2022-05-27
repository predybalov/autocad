"""
Изменение нумерации объекта чертежа, имеющего формат ПРЕФИКС ЗНАЧЕНИЕ [СУФФИКС]
Префикс обязателен, суффикс может отсутствовать
Нумерация изменяется в соответствии с указнным шагов в указанном диапазоне
"""

import tkinter as tk
import win32com.client
import time
from pyautocad import Autocad


def renum():
    global prefix
    global suffix
    global start_num
    global stop_num
    global increment
    prefix = prefix_entry.get()
    suffix = suffix_entry.get()
    start_num = int(start_num_entry.get())
    stop_num = int(stop_num_entry.get())
    increment = int(increment_entry.get())
    prefix_length = len(prefix)
    suffix_length = len(suffix)

    print()
    print('*' * 50)
    print('Начинаю переименование...')
    print('*' * 50)
    print()

    counter = 0
    start_time = time.time()

    acad = win32com.client.Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument  # Document object
    # Iterate trough all objects (entities) in the currently opened drawing
    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference':
            for attrib in entity.GetAttributes():
                # first condition - latin alphabet, second - russian alphabet
                if suffix_length >= 1:
                    if attrib.TextString[:prefix_length] == prefix and \
                            start_num <= int((attrib.TextString[prefix_length:])[:-suffix_length]) <= stop_num:
                        current_num = int((attrib.TextString[prefix_length:])[:-suffix_length])
                        modified_num = str(current_num + increment)
                        modtext = prefix + modified_num + suffix
                        print(attrib.TextString, '->', modtext)
                        attrib.TextString = modtext
                        counter += 1
                if suffix_length == 0:
                    if attrib.TextString[:prefix_length] == prefix:
                        current_num = int(attrib.TextString[prefix_length:])
                        modified_num = str(current_num + increment)
                        modtext = prefix + modified_num
                        print(attrib.TextString, '->', modtext)
                        attrib.TextString = modtext
                        counter += 1
                else:
                    exit(1)
            if (stop_num - start_num + 1) == counter:
                break

    end_time = time.time()
    program_runtime = int(end_time - start_time)

    print()
    acad = Autocad(create_if_not_exists=False)
    print('*' * 50)
    if counter % 10 == 1:
        print(f'Изменён {counter} элемент за {program_runtime} с.')
    elif counter % 10 == 2 or counter % 10 == 3 or counter % 10 == 4:
        print(f'Изменено {counter} элемента за {program_runtime} с.')
    else:
        print(f'Изменено {counter} элементов за {program_runtime} с.')
    print('в файле:', acad.doc.Name)
    print('*' * 50)


counter = ''
prefix = ''
suffix = ''
start_num = ''
stop_num = ''
increment = ''

root = tk.Tk()
width = 290
height = 200
back_color = '#505050'
root.title('Перенумератор 4000')
# Window dimensions 'width x height' and offset from the left top corner
root.geometry(f'{width}x{height}+{int((1920 - width) / 2)}+{int((1080 - height) / 2)}')
root.resizable(False, False)

label_prefix = tk.Label(root, text='Введите префикс')
label_suffix = tk.Label(root, text='Введите суффикс')
label_start_num = tk.Label(root, text='Введите начальное значение')
label_stop_num = tk.Label(root, text='Введите последнее значение')
label_increment = tk.Label(root, text='Введите шаг (+/-)')

prefix_entry = tk.Entry(root)
suffix_entry = tk.Entry(root)
start_num_entry = tk.Entry(root)
stop_num_entry = tk.Entry(root)
increment_entry = tk.Entry(root)

calc_button = tk.Button(root, text='Поехали!', command=renum)

label_prefix.grid(row=0, column=0)
label_suffix.grid(row=1, column=0)
label_start_num.grid(row=2, column=0)
label_stop_num.grid(row=3, column=0)
label_increment.grid(row=4, column=0)

prefix_entry.grid(row=0, column=1)
suffix_entry.grid(row=1, column=1)
start_num_entry.grid(row=2, column=1)
stop_num_entry.grid(row=3, column=1)
increment_entry.grid(row=4, column=1)

calc_button.grid(row=6, column=0, columnspan=2, stick='we')

root.mainloop()
