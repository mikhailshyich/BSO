import os
import pandas as pd
import tkinter as tk

from tkinter import *
from tkinter import ttk
from tkinter import filedialog

# Переменные
list_mark = []
search_list = []
search_res = []
count = 0
range_count = 0
# file = ''

# Настройка окна
root = tk.Tk()
root.title("Поиск кодов в файле Excel")
# root.geometry("500x600")
# root.resizable(False, False)
w = root.winfo_screenwidth()
h = root.winfo_screenheight()
w = w // 2  # середина экрана
h = h // 2
w = w - 450  # смещение от середины
h = h - 530
root.geometry('900x1000+{}+{}'.format(w, h))
font_size = 15


# root.attributes("-toolwindow", True)


def open_file():
    try:
        global list_mark
        filepath = filedialog.askopenfilename(filetypes=(("Файл .xls", " .xls"), ("Все файлы", "*")))
        file = filepath.title()
        file = pd.read_excel(f'{file}')
        list_mark = file['БСО'].tolist()
        list_mark.sort()
        # list_mark = file.iloc[:,0].tolist()
        print("Коды: {0}".format(list_mark))
        print("Файл прочитан. Кодов в файле - {0}".format(len(list_mark)))
        lbl_file["text"] = f"Кодов в файле: {len(list_mark)}"
        add_btn["state"] = "enabled"
        range_bso()
    except:
        print("Что-то пошло не так")


def range_bso():
    global list_mark
    search_bso = []
    list_bso = []
    result = []
    f_number = 1
    s_number = 2
    diap_count = 0
    for item in list_mark:
        list_bso.append(item)
    while len(list_bso) > 0:
        first_code = list_bso[0]
        chast_coda = first_code[:8]
        for num_bso in range(len(list_bso)):
            if chast_coda in list_bso[num_bso]:
                search_bso.append(list_bso[num_bso])
                # index_bso.append(num_bso)
        # print("Часть кодов БСО {0}".format(search_bso))
        first = search_bso[0]
        last = search_bso[-1]
        all_codes = len(search_bso)
        print(
            "Первый код диапазона - {0}. Последний код диапазона - {1}. Всего кодов в диапазоне {2}".format(first, last,
                                                                                                            all_codes))
        rangeBso = ttk.Label(mark_frame,
                             text=f"{f_number} - {first} | {s_number} - {last}. Всего кодов в диапазоне {all_codes}".format(
                                 f_number, s_number, first, last, all_codes), font=font_size)
        rangeBso.pack(anchor=NW)
        # print("Индексы {0}".format(index_bso))
        for bso in list_bso:
            if bso not in search_bso:
                result.append(bso)
        list_bso = []
        for i in result:
            list_bso.append(i)
        # print("Очищенный список - {0}".format(list_bso))
        search_bso = []
        result = []
        f_number += 2
        s_number += 2
        diap_count += 1
    diap_file["text"] = f"Количество диапазонов в файле - {diap_count}"
    # print("Рабочий список после удаления {0}".format(list_bso))
    # print("Обрезанный код для поиска {}".format(first_code[:8]))
    # print("До удаления {0}".format(list_bso[:10]))
    # # deleted_item = list_bso.pop(0)
    # print("После удаления {0}".format(list_bso[:10]))
    # print("Удалённый {0}".format(deleted_item))


def add_bso():
    try:
        global search_list
        search_seria = seria.get()
        search_na4alo = int(na4alo.get())
        search_konec = int(konec.get())
        if len(search_seria) <= 3:
            while search_na4alo <= search_konec:
                search_result = f"{search_seria}{search_na4alo}"
                search_list.append(search_result)
                search_na4alo = search_na4alo + 1
            rangeCount()
            print("Коды которые ищем: {0}".format(search_list))
            search_btn["state"] = "enabled"
        else:
            print("Серия БСО больше 3-х символов")
    except:
        print("Не заполнены все нужные поля в разделе *Работа с диапазонами кодов*")


def rangeCount():
    global range_count
    range_count += 1
    text_count["text"] = f"Количество добавленных диапазонов - {range_count}"


def searchRes():
    global search_res
    search_res = [x for x in search_list if x not in list_mark]
    text_searchRes["text"] = f"Количество отсутствующих кодов - {len(search_res)}"
    print("Отсутствующие коды: {0}".format(search_res))
    save_btn["state"] = "enabled"


def save_file():
    try:
        title_n = title.get()
        batch_n = batch.get()
        prod_date_n = prod_date.get()
        file_name = f"{title_n}_партия_{batch_n}_от_{prod_date_n}.txt"
        file_bco = open(f"{file_name}", "w+")
        for item in search_res:
            file_bco.write("%s\n" % item)
        file_bco.close()
        dir = os.path.abspath(os.curdir)
        print(f"Файл: {dir}\{title_n}_партия_{batch_n}_от_{prod_date_n}.txt".format(dir))
        info_save["text"] = f"Файл: {dir}\{title_n}_партия_{batch_n}_от_{prod_date_n}.txt".format(dir)
        null_btn["state"] = "enabled"
    except:
        print("Дата в неверном формате!")


def clear_all():
    global list_mark
    global search_list
    global search_res
    global count
    global range_count
    list_mark = []
    search_list = []
    search_res = []
    count = 0
    range_count = 0
    text_count["text"] = "Количество добавленных диапазонов вручную -"
    lbl_file["text"] = "Кодов в файле excel -"
    text_searchRes["text"] = "Количество отсутствующих кодов в документе маркировки -"
    info_save["text"] = "Файл .txt сохраняется в ту же директорию, где находится файл .exe"
    title.delete(0, END)
    batch.delete(0, END)
    prod_date.delete(0, END)
    seria.delete(0, END)
    na4alo.delete(0, END)
    konec.delete(0, END)
    add_btn["state"] = "disabled"
    search_btn["state"] = "disabled"
    save_btn["state"] = "disabled"
    for widget in mark_frame.winfo_children():
        widget.destroy()
    diap_file["text"] = "Количество диапазонов в файле -"
    print("Очистка успешна!")


dwnld_frame = ttk.Frame(borderwidth=1, relief=SOLID, padding=5)
downl_btn = ttk.Button(dwnld_frame, text="Загрузить документ маркировки формата .xls", command=open_file,
                       cursor="hand2")
downl_btn.pack(anchor="n", fill=X, expand=True, ipadx=10, ipady=10)  # размещаем кнопку по центру окна
dwnld_frame.pack(fill=X)

# Содержимое окна номенклатура
nomenkl_frame = ttk.Frame(borderwidth=1, relief=SOLID, padding=5)
name_nomenklFrame = ttk.Label(text="Информация для файла .txt", foreground="#E0FFFF", background="#00CED1",
                              font=font_size)
name_title = ttk.Label(nomenkl_frame, text="Номенклатура", font=font_size)
name_batch = ttk.Label(nomenkl_frame, text="Номер партии", font=font_size)
name_prodDate = ttk.Label(nomenkl_frame, text="Дата производства", font=font_size)
title = ttk.Entry(nomenkl_frame)
batch = ttk.Entry(nomenkl_frame)
prod_date = ttk.Entry(nomenkl_frame)

name_nomenklFrame.pack(anchor=NW, fill=X)
name_title.pack(anchor=NW)
title.pack(anchor=NW, fill=X)
name_batch.pack(anchor=NW)
batch.pack(anchor=NW, fill=X)
name_prodDate.pack(anchor=NW)
prod_date.pack(anchor=NW, fill=X)
nomenkl_frame.pack(anchor=NW, fill=X)

# Содержимое окна БСО
bso_frame = ttk.Frame(borderwidth=1, relief=SOLID, padding=5)
name_bsoFrame = ttk.Label(text="Работа с диапазонами кодов", foreground="#E0FFFF", background="#00CED1", font=font_size)
text_seria = ttk.Label(bso_frame, text="Серия БСО", font=font_size)
text_na4alo = ttk.Label(bso_frame, text="Первый код диапазона", font=font_size)
text_konec = ttk.Label(bso_frame, text="Последний код диапазона", font=font_size)
seria = ttk.Entry(bso_frame)
na4alo = ttk.Entry(bso_frame)
konec = ttk.Entry(bso_frame)

add_btn = ttk.Button(bso_frame, text="Добавить", command=add_bso, state="disabled", cursor="plus")
search_btn = ttk.Button(bso_frame, text="Поиск", command=searchRes, state="disabled", cursor="hand2")

name_bsoFrame.pack(anchor=NW, fill=X)
text_seria.pack(anchor=NW)
seria.pack(anchor=NW, fill=X)
text_na4alo.pack(anchor=NW)
na4alo.pack(anchor=NW, fill=X)
text_konec.pack(anchor=NW)
konec.pack(anchor=NW, fill=X)
add_btn.pack(side=LEFT, anchor=NW, expand=True, fill=X, ipadx=10, ipady=10)
search_btn.pack(side=LEFT, anchor=NE, expand=True, fill=X, ipadx=10, ipady=10)
bso_frame.pack(anchor=NW, fill=X)

# Содержимое окна Информация
info_frame = ttk.Frame(padding=5)
mark_frame = ttk.Frame(padding=5)
name_infoFrame = ttk.Label(text="Информация", foreground="#E0FFFF", background="#00CED1", font=font_size)
lbl_file = ttk.Label(info_frame, text="Кодов в файле excel -", font=font_size)
diap_file = ttk.Label(info_frame, text="Количество диапазонов в файле -", font=font_size)
line = ttk.Label(info_frame, text="____________________________________________________________________________",
                 font=font_size)
text_count = ttk.Label(info_frame, text="Количество добавленных диапазонов вручную -", font=font_size)
text_searchRes = ttk.Label(info_frame, text="Количество отсутствующих кодов в документе маркировки -", font=font_size)

name_infoFrame.pack(anchor=NW, fill=X)
mark_frame.pack(anchor=NW, fill=X)
lbl_file.pack(anchor=NW)
diap_file.pack(anchor=NW)
line.pack(anchor=NW)
text_count.pack(anchor=NW)
text_searchRes.pack(anchor=NW)
info_frame.pack(anchor=NW, fill=X)

futter = ttk.Frame(borderwidth=1, relief=SOLID, padding=5)
futter.pack(anchor=SW, fill=X)
info_save = ttk.Label(futter, text="Файл .txt сохраняется в ту же директорию, где находится файл .exe", font=font_size)
info_save.pack(anchor=NW)
save_btn = ttk.Button(futter, text="Сохранить коды в файл", command=save_file, state="disabled", cursor="pencil")
null_btn = ttk.Button(futter, text="Очистить всё", command=clear_all, cursor="exchange")
save_btn.pack(anchor="s", fill=X, expand=True, ipadx=10, ipady=10)  # размещаем кнопку по центру окна
null_btn.pack(anchor="s", fill=X, expand=True, ipadx=5, ipady=5)  # размещаем кнопку по центру окна

# Запуск окна
root.mainloop()
