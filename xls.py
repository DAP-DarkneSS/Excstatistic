#!/usr/bin/env python
# -*- coding: utf-8 -*-
#Указывается язык и кодировка.

from xlrd import *
#Загружается модуль работы с таблицами Excel *.xls.

def xls(filename, t_sheet, t_column):
#Функция принимает строковые значения имени файла, номера листа и колонки.

    matrix = [""]
#Создаётся ряд с пустой строкой (первое значение).

    sheet_number = int(t_sheet)
#Номер листа преобразуется в целое число.

    if len(t_column) == 1:
#Проверка количества символом в номере колонки.

        column_number = (ord(t_column) - 65)
#Номер колонки преобразуется в число в соответствии с кодовой таблицей.

        if (-1 < column_number < 26):
#Проверка, является ли номер колонки буквой от A до Z.

            book = open_workbook(filename)
#Открывается файл.

            sheet = book.sheet_by_index(sheet_number - 1)
#Открывается нужный лист.

            for i in xrange(sheet.nrows):
                cell = sheet.cell_value(rowx=i, colx=column_number)
                matrix.append(cell)
#Значения считываются из колонки и записываются в ряд.

        else:
            matrix.append(u"Ошибка: введите букву столбца от A до Z!")
    else:
        matrix.append(u"Ошибка: введите букву столбца от A до Z!")
#Ошибка возвращается, когда номер колонки не является буквой от A до Z.

        
    return(matrix)
#Функция возвращает ряд.