#!/usr/bin/env python
# -*- coding: utf-8 -*-
#Указывается язык и кодировка.

from Tkinter import *
from tkFileDialog import *
from body import excstatistic
#Загружается основной модуль, графическая библиотека и модуль выбора файла.

root=Tk()
root.title(u"Excstatistic")
root.resizable(False, False)
#Создаётся окно приложения, задаётся заголовок.
#Нельзя изменять размер окна.

label_title = Label(root, text = u"Количество упоминаний элементов")
label_title.pack()
#Надпись с описанием программы.

frame_sheet = Frame(root)
frame_sheet.pack(fill=BOTH)
frame_column = Frame(root)
frame_column.pack(fill=BOTH)
#Создаётся три рамки для виджетов. Рамки растягиваются по ширине окна.

label_sheet = Label(frame_sheet, text = u"Анализируемый лист:")
label_sheet.pack(side=LEFT)
entry_sheet = Entry(frame_sheet, width=3)
entry_sheet.pack(side=RIGHT)
label_column = Label(frame_column, text = u"Анализируемый столбец:")
label_column.pack(side=LEFT)
entry_column = Entry(frame_column, width=3)
entry_column.pack(side=RIGHT)

entry_sheet.insert(END, u"1")
entry_column.insert(END, u"A")

#Надписи, описывающие вводимые значения, выровнены по левому краю.
#Элементы для ввода значений шириной в 3 знака выровнены по правому краю.

filename = ""
def button_fopen():
    global filename
    filename = askopenfilename()

def button_fmake():
#Функция для кнопки. Записывается без аргументов!

    text_out.delete(1.0, END)
    text = excstatistic(filename, entry_sheet.get(), entry_column.get())
    text_out.insert(END, text)
#Передача внешней функции всех параметров. Получение теста и передача его в поле.
    
frame_buttonz = Frame(root)
frame_buttonz.pack(fill=BOTH)
button_open = Button(frame_buttonz, text=u"Открыть файл", command=button_fopen)
button_open.pack(side=LEFT)
button_make = Button(frame_buttonz, text=u"Анализировать", command=button_fmake)
button_make.pack(side=LEFT)
button_exit = Button(frame_buttonz, text=u"Выход", command=root.destroy)
button_exit.pack(side=RIGHT)
#Рамка для кнопок. Кнопки выбора файла, запуска анализа и выхода из приложения.

text_out=Text(root, height=1, width=11)
text_out.pack(fill=BOTH)
#Текстовый виджет растянут по ширине окна приложения.

root.mainloop()
#Окончание текста приложения.