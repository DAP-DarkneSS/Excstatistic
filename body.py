#!/usr/bin/env python
# -*- coding: utf-8 -*-
#Указывается язык и кодировка.

from xls import xls
from xlsx import xlsx
from os import system
from sys import platform
#Загружаются модули импорта таблиц Excel *.xls,
#модуль определения операционной системы и работы с терминалом.

def excstatistic(filename, t_sheet, t_column):
#Функция принимает строковые значения имени файла, номера листа и колонки.

    text = u""
    name = []
    name_count = []
    matrix = []
#Создаётся пустая строка и ряды для импортированных значений
#и проанализированных значений и количества их упоминания.

    if filename.endswith(".xlsx"):
        matrix = xlsx(filename, t_sheet, t_column)
    elif filename.endswith(".xls"):
        matrix = xls(filename, t_sheet, t_column)
    else:
        matrix.append(u"Ошибка: откройте файл таблицы Excel!")
#Выбор необходимого модуля, в случае неправильного файла возвращается ошибка.

    if not (matrix[0]).startswith(u"Ошибка: "):
#Проверка на ошибки.

        for i in xrange(len(matrix)):
            if matrix[i] != "":
                j = -1
                try:
                    j = name.index(matrix[i])
                except ValueError:
                    pass
                if (-1 < j < len(name)):
                    name_count[j] += 1
#Если элемент их импортированного ряда нахоидтся в ряде с проанализированными значениями,
#то соотвествующий счётчик увеличивается на единицу.

                else:
                    name.append(matrix[i])
                    name_count.append(1)
#Иначе импортированное значение добавляется к ряду с проанализированными значениями
#и создаётся соотвествующий счётчик.

        if platform.startswith("win"):
            system('cls')
        else:
            system('clear')
        print(u"Результаты:")
#Очистка консоли перед выводом результатов и объявление вывода.

        for i in xrange(len(name)):
            print(u"\n" + str(name_count[i]) + u" @ ")
            print(name[i])
        text = u"Результаты получены и выведены в консоли."
#Вывод значений, заполнение строки предупреждением.

    else:
        text = matrix[0]
#Если импортированный ряд начинается с сообщения об ошибке,
#то оно передаётся в стровое значение.

    return(text)
#Функция возвращает строку.