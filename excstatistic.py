#!/usr/bin/env python
# -*- coding: utf-8 -*-
#Указывается язык и кодировка.

from Tkinter import *
from tkFileDialog import *
from body import excstatistic
#Загружается основной модуль, графическая библиотека и модуль выбора файла.

class MyEntry:
#Класс для улучшения читаемости кода однотипных элементов для ввода параметров.
    def __init__(self, place_class, string_class, string_2_class):
#При создании принимается место прикрепления виджета и строковое значение для надписи.
        self.frame_class = Frame(place_class)
        self.frame_class.pack(side = TOP, fill = BOTH)
#Внутри – рамка для виджетов, растягивается по ширине окна.
        self.label_class = Label(self.frame_class, text = string_class)
        self.label_class.pack(side = LEFT)
#В ней – надписи для описания вводимых значений выровнены по левому краю.
        self.entry_class = Entry(self.frame_class, width = 3)
        self.entry_class.pack(side = RIGHT)
#И элементы для ввода значений шириной в 3 знаков выровнены по правому краю.
        self.entry_class.insert(END, string_2_class)
#Вставка значения по умолчанию.
    def get(self):
        return(self.entry_class.get())
#Метод .get() передаётся от элемента для ввода объекту описываемого класса.

class MyButton:
#Класс для улучшения читаемости кода однотипных элементов кнопок.
    def __init__(self, place_class, string_class, command_class):
#При создании принимается место прикрепления виджета, строковое значение для надписи
#и строковое для установления команды при нажатии.
        self.button_class = Button(place_class, text = string_class, command = command_class)
        self.button_class.pack(side = LEFT)
#Кнопка прикрепляется к левому краю.

def button_fopen():
#Функция для кнопки открытия файла.
    global filename
#Переменная объявляется глобальной, чтобы вывести значение за функцию.
    filename = askopenfilename()
#Собственно открытие файла.

def button_fmake():
#Функция для кнопки. Записывается без аргументов!
    text_info = excstatistic(filename, entry_sheet.get(), entry_column.get())
    text_out.config(text = text_info)
#Передача внешней функции всех параметров. Получение теста и передача его в надпись.

filename = ""
#Создаётся пустая переменная для имени файла.

root=Tk()
#Создаётся окно приложения.
root.title(u"Excstatistic")
#Задаётся заголовок.
root.resizable(False, False)
#Нельзя изменять размер окна.

label_title = Label(root, text = u"Количество упоминаний элементов\nв столбце Excel")
label_title.pack()
#Надпись с описанием программы.

entry_sheet = MyEntry(root, u"Анализируемый лист:", u"1")
entry_column = MyEntry(root, u"Анализируемый столбец:", u"A")
#Создаётся необходимое количество объектов класса элементов ввода.

frame_buttonz = Frame(root)
frame_buttonz.pack(side = TOP, fill = BOTH)
#Рамка для кнопок.
button_open = MyButton(frame_buttonz, u"Открыть файл", button_fopen)
button_make = MyButton(frame_buttonz, u"Анализировать", button_fmake)
button_exit = MyButton(frame_buttonz, u"Выход", root.destroy)
#Кнопки выбора файла, запуска анализа и выхода из приложения.

text_out = Label(root)
text_out.pack(side = LEFT, fill = BOTH)
#Информационная надпись растянута по ширине окна приложения, прикреплена к левому краю.

root.mainloop()
#Окончание текста приложения.
