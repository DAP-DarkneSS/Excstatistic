#!/usr/bin/env python
# -*- coding: utf-8 -*-

#This code is released under
#GNU LESSER GENERAL PUBLIC LICENSE
#http://www.gnu.org/copyleft/lesser.html
#Author: DarkneSS at https://github.com/DAP-DarkneSS

from Tkinter import *
from ScrolledText import *
from tkFileDialog import *
from xlrd import *
#Загружается модуль импорта таблиц Excel, графическая библиотека
#и модуль, содержащий текстовый виджет с полосой прокрутки.

root=Tk()
root.title(u"Статистика по таблице")
#Создаётся окно приложения, задаётся заголовок.

label_title = Label(root, text = u"Статистика по выбранному столбцу таблицы")
label_title.pack()
#Надпись, дублирующая заголовок.

frame_sheet = Frame(root)
frame_sheet.pack(fill=BOTH)
frame_column = Frame(root)
frame_column.pack(fill=BOTH)
##frame_finish = Frame(root)
##frame_finish.pack(fill=BOTH)
#Создаётся три рамки для виджетов. Рамки растягиваются по ширине окна.

label_sheet = Label(frame_sheet, text = u"Анализируемый лист:")
label_sheet.pack(side=LEFT)
entry_sheet = Entry(frame_sheet, width=3)
entry_sheet.pack(side=RIGHT)
label_column = Label(frame_column, text = u"Анализируемый столбец:")
label_column.pack(side=LEFT)
entry_column = Entry(frame_column, width=3)
entry_column.pack(side=RIGHT)
##label_finish = Label(frame_finish, text = u"Количество пустых строк для завершения:")
##label_finish.pack(side=LEFT)
##entry_finish = Entry(frame_finish, width=3)
##entry_finish.pack(side=RIGHT)

entry_sheet.insert(END, u"1")
entry_column.insert(END, u"A")
##entry_finish.insert(END, "7")

#Надписи, описывающие вводимые значения, выровнены по левому краю.
#Элементы для ввода значений шириной в 15 знаков выровнены по правому краю.

def column_error():
    text_out.insert(END, u"Ошибка: введите букву столбца от A до Z!")

def button_fopen():
    global book
    filename = askopenfilename()
    book = open_workbook(filename)

def button_fmake():
#Функция для кнопки. Записывается без аргументов!

    text_out.delete(1.0, END)
    sheet_nunmber = int(entry_sheet.get())
    column_letter = entry_column.get()
    if len(column_letter) == 1:
        column_nunmber = (ord(column_letter) - 65)
        if not (-1 < column_nunmber < 26):
            column_error()
    else:
        column_error()

    sheet = book.sheet_by_index(sheet_nunmber - 1)

    name = []
    name_count = []


    for i in xrange(sheet.nrows):
        cell = sheet.cell_value(rowx=i, colx=column_nunmber)
        if cell != "":
            j = -1
            try:
                j = name.index(cell)
            except ValueError:
                pass
            if (-1 < j < len(name)):
                name_count[j] += 1
            else:
                name.append(cell)
                name_count.append(1)

    for i in xrange(len(name)):
##        text_out.insert(END, (str(name_count[i])), u" @ ", name[i], u"\n")
        text_out.insert(END, (str(name_count[i])))
        text_out.insert(END, u" @ ")
        text_out.insert(END, name[i])
        text_out.insert(END, u"\n")
##    print(filename.get())
##    text_out.insert(END, str(filename))
##    mini = float((entry_mini.get()).replace(",", "."))
##    maxi = float((entry_maxi.get()).replace(",", "."))
#Очистка текстового поля. Импорт минимума и максимума,
#замена запятых на точки, если необходимо.

##    if vcheck_punctuation.get():
##        for i in xrange((int(entry_n.get()) - 1)):
##            text_out.insert(END, str(uniform(mini, maxi))+"\n")
##        text_out.insert(END, str(uniform(mini, maxi)))
#Если не требуется вывод чисел с запятыми,
#текстовое поле заполняется случайными числами,
#каждое с новой строки. Последнее значение вставляется без перевода строки.

##    else:
##        for i in xrange((int(entry_n.get()) - 1)):
##            text_out.insert(END, str(uniform(mini, maxi)).replace(".", ",")+"\n")
##        text_out.insert(END, str(uniform(mini, maxi)).replace(".", ","))
#Иначе в аналогичных операциях точки заменяются на запятые.

##    if vcheck_copy.get():
##        root.clipboard_clear()
##        root.clipboard_append(text_out.get(1.0, END))
#Если не указано иное, очищается буфер обмена, копируются полученные значения.

frame_buttonz = Frame(root)
frame_buttonz.pack(fill=BOTH)
button_open = Button(frame_buttonz, text=u"Открыть файл", command=button_fopen)
button_open.pack(side=LEFT)
button_make = Button(frame_buttonz, text=u"Анализировать", command=button_fmake)
button_make.pack(side=LEFT)
button_exit = Button(frame_buttonz, text=u"Выход", command=root.destroy)
button_exit.pack(side=RIGHT)
#Рамка для кнопок. Кнопка генерирования и выхода из приложения.

##frame_checkz = Frame(root)
##frame_checkz.pack(fill=BOTH)
##frame_copy = Frame(frame_checkz)
##frame_copy.pack(fill=BOTH)
##frame_punctuation = Frame(frame_checkz)
##frame_punctuation.pack(fill=BOTH)
##vcheck_copy=IntVar()
##vcheck_punctuation=IntVar()
#Рамки для чекбоксов и создание переменных для значений чекбоксов.

##check_copy=Checkbutton(frame_copy,text=u'Автоматически копировать результат',variable=vcheck_copy,onvalue=1,offvalue=0)
##check_copy.pack(side=LEFT)
##check_copy.select()
#Чекбокс для включения/выключения автоматического копирования значений.
#Активен по умолчанию и означает "копировать".

##check_punctuation=Checkbutton(frame_punctuation,text=u'Выводить с точками (по умолчанию - запятые)',variable=vcheck_punctuation,onvalue=1,offvalue=0)
##check_punctuation.pack(side=LEFT)
#Чекбокс для переключения между точками и запятыми.
#Неактивен по умолчанию и определяет вывод с запятыми.

text_out=ScrolledText(root, height=15, width=15)
text_out.pack(fill=BOTH)
#Текстовый виджет с полосой прокрутки растянут по ширине окна приложения.

root.mainloop()
#Окончание текста приложения.
